﻿using System;
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
    /// Логика взаимодействия для FormEditEmployee.xaml
    /// </summary>
    public partial class FormEditEmployee : Window
    {
        string ID = "";
        public FormEditEmployee(string ID = "")
        {
            this.ID = ID;
            InitializeComponent();
            DateTime dateTimeNow = DateTime.Now;
            DateTime dateTimeMin = dateTimeNow.AddYears(-52);
            dpDate.DisplayDateStart = dateTimeMin;
            dpDate.DisplayDateEnd = DateTime.Now.AddYears(-18);
            dpDate.Text = dpDate.DisplayDateEnd.ToString();

            DataTable Table = new DataTable();
            SqlDataAdapter Adapter;
            if (ID != null)
            {
                Adapter = new SqlDataAdapter("SELECT Full_name, Date_of_birth, Passport, Phone_number, Login, Password FROM Employee WHERE Employee_num = " + ID, DataBank.sqlConnection);
                Adapter.Fill(Table);
                tbFIO.Text = Table.Rows[0][0].ToString();
                dpDate.Text = Table.Rows[0][1].ToString();
                tbPassport.Text = Table.Rows[0][2].ToString();
                tbPhone.Text = Table.Rows[0][3].ToString();
                tbLogin.Text = Table.Rows[0][4].ToString();
                tbPassword.Text = Table.Rows[0][5].ToString();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ID != null)
                {
                    DataBank.sqlConnection.Open();
                    SqlCommand commandabonent = new SqlCommand("UPDATE Employee SET Full_name = '" + tbFIO.Text + "', Date_of_birth = '" + dpDate.Text + "', Phone_number = '" + tbPhone.Text + "', Passport = '" + tbPassport.Text + "', Login = '" + tbLogin.Text + "', Password = '" + tbPassword.Text + "' WHERE Employee_num = " + ID, DataBank.sqlConnection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные изменены!");
                    DataBank.sqlConnection.Close();
                }
                else
                {
                    DataBank.sqlConnection.Open();
                    SqlCommand commandabonent = new SqlCommand($"Insert Into Employee (Full_name, Date_of_birth, Phone_number, Passport, Login, Password, Post_num) Values ('{tbFIO.Text}', '{dpDate.Text}', '{tbPhone.Text}', '{tbPassport.Text}', '{tbLogin.Text}', '{tbPassport.Text}', 2)", DataBank.sqlConnection);
                    commandabonent.ExecuteNonQuery();

                    MessageBox.Show("Данные добавлены!");
                    DataBank.sqlConnection.Close();
                }
            }
            catch { MessageBox.Show("Данные введены неверно!", "Ошибка"); }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            DataBank.formAdmin.LoadDataEmployee();
            Close();
        }
    }
}
