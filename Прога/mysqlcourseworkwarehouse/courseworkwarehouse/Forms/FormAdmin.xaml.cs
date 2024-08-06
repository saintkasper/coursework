using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using Window = Microsoft.Office.Interop.Excel.Window;
using DataTable = System.Data.DataTable;
using Microsoft.Win32;
using MySql.Data.MySqlClient;

namespace courseworkwarehouse.Forms
{
    /// <summary>
    /// Логика взаимодействия для FormAdmin.xaml
    /// </summary>
    public partial class FormAdmin : Window
    {
        string FilePath = "";
        string data = "";
        string AscDesc = "";
        string Filter = "";
        int k = 0;

        public FormAdmin()
        {
            InitializeComponent();
            DataBank.formAdmin = this;
            SMassegeAdmin.Content = $"Добро пожаловать {DataBank.userFIO}!";
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
                SnackAdmin.IsActive = false;
            }
            else
            {
                SnackAdmin.IsActive = true;
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


        #region Employee
        public void LoadDataEmployee()
        {
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"select Full_name as ФИО, Date_of_birth as 'Дата Рождения', Passport as 'Паспортные Данные', Phone_number as 'Номер Телефона', Login as Логин, Password as Пароль, Employee_num as 'Код' From Employee WHERE Full_name Like '%{tbSearch.Text}%' {AscDesc}", MySql.connection);
            adapter.Fill(table);
            dg.ItemsSource = table.DefaultView;
        }

        private void btnExitEmployee_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)dg.SelectedItems[0];
            FormEditEmployee editEmployee = new FormEditEmployee(row[6].ToString());
            editEmployee.Show();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            FormEditEmployee editEmployee = new FormEditEmployee(null);
            editEmployee.Show();
        }

        private void tbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            LoadDataEmployee();
        }

        private void rbWithout_Checked(object sender, RoutedEventArgs e)
        {
            if (rbWithout.IsChecked == true)
            {
                AscDesc = "";
                LoadDataEmployee();
            }
        }

        private void rbAsc_Checked(object sender, RoutedEventArgs e)
        {
            if (rbAsc.IsChecked == true)
            {
                AscDesc = "ORDER BY Full_name ASC";
                LoadDataEmployee();
            }
        }

        private void rbDesc_Checked(object sender, RoutedEventArgs e)
        {
            if (rbDesc.IsChecked == true)
            {
                AscDesc = "ORDER BY Full_name DESC";
                LoadDataEmployee();
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            DataRowView rows = (DataRowView)dg.SelectedItems[0];
            if (DataBank.employeenum == rows[6].ToString())
            {
                MessageBox.Show("Вы не можете удалить самого себя!", "Ошибка!");
            }
            else
            {
                if (dg.SelectedItems.Count == 0)
                {
                    MessageBox.Show("Выберите сотрудника для удаления");
                }
                else
                {

                    MySql.connection.Open();

                    DataRowView row = (DataRowView)dg.SelectedItems[0];

                    MySqlCommand commandcity = new MySqlCommand("DELETE FROM Employee WHERE Employee_num=" + row["Код Сотрудника"].ToString(), MySql.connection);
                    commandcity.ExecuteNonQuery();

                    MySql.connection.Close();

                    MessageBox.Show("Данные удалены!");
                    LoadDataEmployee();
                }
            }
        }

        private void tiEmployee_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataEmployee();
        }
        #endregion

        #region Product
        public void LoadDataProduct()
        {
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"select Product.Product_article as 'Артикул товара', Product.Title as 'Название', Product.Cost as 'Цена', Product.Quantity as 'Количество', Product.Supplier_num as 'Код поставщика', Product.Product_group_num as 'Код товарной группы', Supplier.Title as 'Поставщик', `product group`.Title as 'Товарная группа', Photo From Product inner join Supplier on Product.Supplier_num = Supplier.Supplier_num inner join `product group` on `product group`.Product_group_num = Product.Product_group_num WHERE Product.Title Like '%{tbSearchProduct.Text}%' {Filter} {AscDesc}", MySql.connection);
            adapter.Fill(table);
            dgProduct.ItemsSource = table.DefaultView;

            if (DataBank.countGroups != 2)
                cbRefresh();

            try
            {
                dgProduct.Columns[10].MaxWidth = 0;
            }
            catch
            { }
        }

        private void btnExitProduct_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }


        private void cbRefresh()
        {
            int currentCount = cbProduct.Items.Count;

            if (cbProduct.Items.Count != 0)
            {
                if (cbProduct.Items.Count != currentCount || DataBank.countGroups == 1)
                {
                    int cbCount = cbProduct.Items.Count - 1;
                    for (int i = 0; i < cbCount; i++)
                    {
                        cbProduct.Items.RemoveAt(cbCount - i);
                    }
                }
                DataTable table = new DataTable();
                MySqlDataAdapter adapter = new MySqlDataAdapter($"select distinct Product.Product_group_num, `product group`.Title From Product join `product group` on `product group`.Product_group_num = Product.Product_group_num ", MySql.connection);
                adapter.Fill(table);

                for (int i = 0, c = 1; i < table.Rows.Count; i++, c++)
                {
                    cbProduct.Items.Insert(c, table.Rows[i][1].ToString());
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
                DataBank.countGroups = 1;
                MySql.connection.Open();

                DataRowView row = (DataRowView)dgProduct.SelectedItems[0];

                MySqlCommand commandcity = new MySqlCommand("DELETE FROM Product WHERE Product_article = " + row["Артикул товара"].ToString(), MySql.connection);
                commandcity.ExecuteNonQuery();

                MySql.connection.Close();

                MessageBox.Show("Данные удалены!");
                LoadDataProduct();
            }
        }

        private void btnEditProduct_Click(object sender, RoutedEventArgs e)
        {
            DataRowView row = (DataRowView)dgProduct.SelectedItems[0];
            FormEditProduct editProduct = new FormEditProduct(row[0].ToString());
            editProduct.Show();
        }


        public void tiProduct_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataProduct();
        }

        private void cbProduct_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (tbSearchProduct != null)
            {
                if (cbProduct.SelectedIndex != 0)
                {
                    DataBank.countGroups = 2;
                    Filter = $" And `product group`.Title = '{cbProduct.SelectedItem}'";
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

        private void tiSupplier_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataSupplier();
        }

        private void tbSearchSupplier_TextChanged(object sender, TextChangedEventArgs e)
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

        private void btnDeleteSupplier_Click(object sender, RoutedEventArgs e)
        {
            if (dgSupplier.SelectedItems.Count == 0)
            {
                MessageBox.Show("Выберите товар для удаления");
            }
            else
            {

                MySql.connection.Open();

                DataRowView row = (DataRowView)dgSupplier.SelectedItems[0];

                MySqlCommand commandcity = new MySqlCommand("DELETE FROM Supplier WHERE Supplier_num = " + row["Код поставщика"].ToString(), MySql.connection);
                commandcity.ExecuteNonQuery();

                MySql.connection.Close();

                MessageBox.Show("Данные удалены!");
                LoadDataSupplier();
            }
        }

        private void btnEditSupplier_Click(object sender, RoutedEventArgs e)
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
                    MessageBox.Show("Выберите товар для удаления");
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

        #region Report
        private void tiReport_Loaded(object sender, RoutedEventArgs e)
        {
            LoadDataReport();
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($"Select `tabular part of the receipt`.Receipt_invoice_num as '№ расходной накладной', Product.Title as 'Наименование товара', `tabular part of the receipt`.Quantity as 'Количество', `tabular part of the receipt`.Cost as 'Цена  (руб.)', `tabular part of the receipt`.Date_of_the_receipt as 'Дата составления', Employee.Full_name as 'ФИО сотрудника' From `tabular part of the receipt` join Product on Product.Product_article = `tabular part of the receipt`.Product_article join Employee on Employee.Employee_num=`tabular part of the receipt`.Employee_num", MySql.connection);
            adapter.Fill(table);
            try
            {
                table = new DataTable();
                adapter = new MySqlDataAdapter($"select MIN(Date_of_the_receipt) from `tabular part of the receipt`", MySql.connection);
                adapter.Fill(table);
                Date1.DisplayDateStart = DateTime.Parse(table.Rows[0][0].ToString());
                Date1.SelectedDate = DateTime.Parse(table.Rows[0][0].ToString());
                Date2.DisplayDateStart = DateTime.Parse(table.Rows[0][0].ToString());
                table = new DataTable();
                adapter = new MySqlDataAdapter($"select MAX(Date_of_the_receipt) from `tabular part of the receipt`", MySql.connection);
                adapter.Fill(table);
                Date2.DisplayDateEnd = DateTime.Parse(table.Rows[0][0].ToString());
                Date2.SelectedDate = DateTime.Parse(table.Rows[0][0].ToString());
                Date1.DisplayDateEnd = DateTime.Parse(table.Rows[0][0].ToString());
            }
            catch { }
        }

        public void LoadDataReport()
        {
            DataTable table = new DataTable();
            MySqlDataAdapter adapter = new MySqlDataAdapter($" Select `tabular part of the receipt`.Receipt_invoice_num as '№ расходной накладной', Product.Title as 'Наименование товара', `tabular part of the receipt`.Quantity as 'Количество', `tabular part of the receipt`.Cost as 'Цена  (руб.)', `tabular part of the receipt`.Date_of_the_receipt as 'Дата составления', Employee.Full_name as 'ФИО сотрудника' From `tabular part of the receipt` join Product on Product.Product_article = `tabular part of the receipt`.Product_article join Employee on Employee.Employee_num=`tabular part of the receipt`.Employee_num {data}", MySql.connection);
            adapter.Fill(table);
            dgReportProducts.ItemsSource = table.DefaultView;
        }

        private void Date1_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            Date2.DisplayDateStart = Date1.SelectedDate;
            if (Date1.Text != "" && Date2.Text != "")
            {
                data = $" And `tabular part of the receipt`.Date_of_the_receipt >= STR_TO_DATE('{Date1.Text}','%d.%m.%Y') And `tabular part of the receipt`.Date_of_the_receipt <= STR_TO_DATE('{Date2.Text}','%d.%m.%Y')";
                LoadDataReport();
            }
        }

        private void Date2_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            Date1.DisplayDateEnd = Date2.SelectedDate;
            if (Date1.Text != "" && Date2.Text != "")
            {
                data = $" And `tabular part of the receipt`.Date_of_the_receipt >= STR_TO_DATE('{Date1.Text}','%d.%m.%Y') And `tabular part of the receipt`.Date_of_the_receipt <= STR_TO_DATE('{Date2.Text}','%d.%m.%Y')";
                LoadDataReport();
            }
        }

        private void btnExitReport_Click(object sender, RoutedEventArgs e)
        {
            DataBank.user = "";
            Authorization authorization = new Authorization();
            Hide();
            authorization.ShowDialog();
        }

        private void btnPrint_Click(object sender, RoutedEventArgs e) // Вывод в эксель
        {
            if (dgReportProducts.ItemsSource != null)
            {
                DataTable reportTable = ((DataView)dgReportProducts.ItemsSource).ToTable();
                int cost = 0;

                for (int i = 0; i < reportTable.Rows.Count; i++)
                {
                    cost += int.Parse(reportTable.Rows[i][3].ToString());
                }

                Excel.Application excel = new Excel.Application();
                Workbook workbook = excel.Workbooks.Open(@"C:\Users\furuh\Documents\RReport.xlsx");
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];


                Range myRange1 = (Range)sheet1.Cells[5, 1];
                myRange1.Value2 = $"Отчет по проданным товарам за период с {Date1.Text} по {Date2.Text}";
                myRange1 = (Range)sheet1.Cells[6, 4];
                myRange1.Value2 += $"{DateTime.Now.ToShortDateString()}";
                Range myRange = null;
                for (int j = 0; j < reportTable.Columns.Count; j++)
                {
                    myRange = (Range)sheet1.Cells[9, j + 1]; // Откуда начинаем
                    sheet1.Cells[9, j + 1].Font.Bold = true; // В этих ячейках ставим жирный шрифт
                    myRange.Value2 = reportTable.Columns[j].ColumnName; // Добавляем названия колонок ([j].ColumnName) в первую строку 
                }

                for (int i = 0; i < reportTable.Rows.Count; i++)
                {
                    for (int j = 0; j < reportTable.Columns.Count; j++)
                    {
                        myRange = (Range)sheet1.Cells[10 + i, j + 1];
                        myRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        myRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        myRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                        myRange.Value2 = reportTable.Rows[i][j].ToString();
                    }
                }
                Range myRange2 = sheet1.get_Range((Range)sheet1.Cells[reportTable.Rows.Count + 10, 2], (Range)sheet1.Cells[reportTable.Rows.Count + 10, 3]);
                myRange2.HorizontalAlignment = XlHAlign.xlHAlignRight;
                myRange2.Merge();
                myRange2.Value2 = "Общая сумма (руб.): ";
                myRange1 = (Range)sheet1.Cells[4][reportTable.Rows.Count + 10];
                myRange1.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                myRange1.Value2 = cost.ToString();
                myRange1 = (Range)sheet1.Cells[1][reportTable.Rows.Count + 12];
                myRange1.Value2 = "Подпись: ________";
                myRange1 = (Range)sheet1.Cells[4][reportTable.Rows.Count + 12];
                myRange1.Value2 = "Печать: ";
                Range myRange6 = sheet1.UsedRange;
                myRange6.Columns.AutoFit();
                excel.Visible = true;
            }
            else MessageBox.Show("Что-то пошло не так, проверьте выбранные даты!", "Ошибка!");

        } 

        
        #endregion

        #region Excel
        dynamic Window.Activate()
        {
            throw new NotImplementedException();
        }

        public dynamic ActivateNext()
        {
            throw new NotImplementedException();
        }

        public dynamic ActivatePrevious()
        {
            throw new NotImplementedException();
        }

        public bool Close(object SaveChanges, object Filename, object RouteWorkbook)
        {
            throw new NotImplementedException();
        }

        public dynamic LargeScroll(object Down, object Up, object ToRight, object ToLeft)
        {
            throw new NotImplementedException();
        }

        public Window NewWindow()
        {
            throw new NotImplementedException();
        }

        public dynamic _PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }

        public dynamic PrintPreview(object EnableChanges)
        {
            throw new NotImplementedException();
        }

        public dynamic ScrollWorkbookTabs(object Sheets, object Position)
        {
            throw new NotImplementedException();
        }

        public dynamic SmallScroll(object Down, object Up, object ToRight, object ToLeft)
        {
            throw new NotImplementedException();
        }

        public int PointsToScreenPixelsX(int Points)
        {
            throw new NotImplementedException();
        }

        public int PointsToScreenPixelsY(int Points)
        {
            throw new NotImplementedException();
        }

        public dynamic RangeFromPoint(int x, int y)
        {
            throw new NotImplementedException();
        }

        public void ScrollIntoView(int Left, int Top, int Width, int Height, object Start)
        {
            throw new NotImplementedException();
        }

        public dynamic PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
        {
            throw new NotImplementedException();
        }

        public Excel.Application Application => throw new NotImplementedException();

        public XlCreator Creator => throw new NotImplementedException();

        dynamic Window.Parent => throw new NotImplementedException();

        public Range ActiveCell => throw new NotImplementedException();

        public Chart ActiveChart => throw new NotImplementedException();

        public Pane ActivePane => throw new NotImplementedException();

        public dynamic ActiveSheet => throw new NotImplementedException();

        public dynamic Caption { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayFormulas { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayGridlines { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayHeadings { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayHorizontalScrollBar { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayOutline { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool _DisplayRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayVerticalScrollBar { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayWorkbookTabs { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayZeros { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool EnableResize { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool FreezePanes { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public int GridlineColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public XlColorIndex GridlineColorIndex { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Index => throw new NotImplementedException();

        public string OnWindow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Panes Panes => throw new NotImplementedException();

        public Range RangeSelection => throw new NotImplementedException();

        public int ScrollColumn { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public int ScrollRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Sheets SelectedSheets => throw new NotImplementedException();

        public dynamic Selection => throw new NotImplementedException();

        public bool Split { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public int SplitColumn { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public double SplitHorizontal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public int SplitRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public double SplitVertical { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public double TabRatio { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public XlWindowType Type => throw new NotImplementedException();

        public double UsableHeight => throw new NotImplementedException();

        public double UsableWidth => throw new NotImplementedException();

        public bool Visible { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public Range VisibleRange => throw new NotImplementedException();

        public int WindowNumber => throw new NotImplementedException();

        XlWindowState Window.WindowState { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public dynamic Zoom { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public XlWindowView View { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public SheetViews SheetViews => throw new NotImplementedException();

        public dynamic ActiveSheetView => throw new NotImplementedException();

        public bool DisplayRuler { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool AutoFilterDateGrouping { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
        public bool DisplayWhitespace { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public int Hwnd => throw new NotImplementedException();


        #endregion



        private void btnSaveBackup_Click(object sender, RoutedEventArgs e)
        {
            MySql.connection.Open();

            MySqlCommand backup = new MySqlCommand("backup database [Office supplies warehouse] to disk = '" + @"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup" + "\\" + "[Office supplies warehouse]" + "-" + DateTime.Now.ToString("dd-MM-yyyy--HH-mm-ss") + ".bak'", MySql.connection);
            backup.ExecuteNonQuery();

            MySql.connection.Close();

            MessageBox.Show("Резервная копия БД создана", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
        }

        private void btnBackupPath_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Database restore";
            dlg.InitialDirectory = @"C:\Program Files\Microsoft SQL Server\MSSQL15.SQLEXPRESS\MSSQL\Backup";
            dlg.Filter = "SQL SERVER database backup files|*.bak";
            dlg.ShowDialog();

            FilePath = dlg.FileName;
        }

        private void btnRecover_Click(object sender, RoutedEventArgs e)
        {
            if (FilePath != "")
            {
                MySql.connection.Open();

                MySqlCommand restore = new MySqlCommand("alter database [Office supplies warehouse] set single_user with rollback immediate use master restore database [Office supplies warehouse] from disk = '" + FilePath + "' with replace alter database [Office supplies warehouse] set multi_user", MySql.connection);
                restore.ExecuteNonQuery();

                MySql.connection.Close();

                FilePath = "";

                MessageBox.Show("Копия БД восстановлена", "", MessageBoxButton.OK, MessageBoxImage.Asterisk);
            }
            else
            {
                MessageBox.Show("Путь к файлу не выбран", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
