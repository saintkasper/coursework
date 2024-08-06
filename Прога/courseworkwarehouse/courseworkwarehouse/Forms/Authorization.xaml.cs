using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading;
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
    /// Логика взаимодействия для Authorization.xaml
    /// </summary>
    public partial class Authorization : Window
    {
        int k = 0;
        public Authorization()
        {
            InitializeComponent();
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
                Snack.IsActive = false;
            }
            else
            {
                Snack.IsActive = true;
                k++;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string id = "";
            Authorization mainWindow = new Authorization();
            if (loginTextBox.Text.Length > 0)
            {
                if (passwordBox.Password.Length > 0)
                {
                    DataTable dt = new DataTable();
                    SqlDataAdapter adapter = new SqlDataAdapter("SELECT Post_num, Employee_num FROM [dbo].[Employee] WHERE [Login] = '" + loginTextBox.Text + "' AND [Password] = '" + passwordBox.Password + "'", DataBank.sqlConnection);
                    adapter.Fill(dt);
                    if (dt.Rows.Count > 0)
                    {
                        if (dt.Rows[0][0].ToString() == "1")
                        {
                            id = dt.Rows[0][1].ToString();
                            DataTable table = new DataTable();
                            adapter = new SqlDataAdapter($"SELECT Full_name, Employee_num FROM [dbo].[Employee] WHERE Employee_num = '{id}' And [Login] = '" + loginTextBox.Text + "' AND [Password] = '" + passwordBox.Password + "'", DataBank.sqlConnection);
                            adapter.Fill(table);
                            DataBank.userFIO = table.Rows[0][0].ToString();
                            DataBank.user = "admin";
                            DataBank.employeenum = table.Rows[0][1].ToString();
                            FormAdmin adminForm = new FormAdmin();
                            Hide();
                            adminForm.ShowDialog();
                        }
                        else
                        {
                            id = dt.Rows[0][1].ToString();
                            DataTable table = new DataTable();
                            adapter = new SqlDataAdapter($"SELECT Full_name, Employee_num FROM [dbo].[Employee] WHERE Employee_num = '{id}' And [Login] = '" + loginTextBox.Text + "' AND [Password] = '" + passwordBox.Password + "'", DataBank.sqlConnection);
                            adapter.Fill(table);
                            DataBank.userFIO = table.Rows[0][0].ToString();
                            DataBank.user = "manager";
                            DataBank.employeenum = table.Rows[0][1].ToString();
                            FormManager formManager = new FormManager();
                            Hide();
                            formManager.ShowDialog();
                        }
                    }
                    else MessageBox.Show("Такого пользователя не существует!");
                }
                else MessageBox.Show("Введите пароль!");
            }
            else MessageBox.Show("Введите логин!");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Application.Current.Shutdown();
        }
    }
}
