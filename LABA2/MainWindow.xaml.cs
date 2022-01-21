using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;

namespace LABA2
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        SqlConnection sqlConnection;
        List<string> Before = new List<string>();
        public MainWindow()
        {
            InitializeComponent();
            
        }

        public int update_sum = 0;
        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {

            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\QQQ\Desktop\LABA2\LABA2\Database1.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            SqlCommand command = new SqlCommand("DROP TABLE InformationSecurityThreats1", sqlConnection);
            await command.ExecuteNonQueryAsync();
            command = new SqlCommand("CREATE TABLE InformationSecurityThreats1 " +
                "([id] INT IDENTITY NOT NULL PRIMARY KEY, " +
                "[ИдентификаторУБИ]            INT            NULL, " +
                "[НаименованиеУБИ]             NVARCHAR (MAX) NULL, " +
                "[Описание]                    NVARCHAR (MAX) NULL, " +
                "[ИсточникУгрозы]              NVARCHAR (MAX) NULL, " +
                "[ОбъектВоздействияУгрозы]     NVARCHAR (MAX) NULL, " +
                "[НарушениеКонфиденциальности] INT            NULL, " +
                "[НарушениеЦелостности]        INT            NULL, " +
                "[НарушениеДоступности]        INT            NULL)", sqlConnection);
            await command.ExecuteNonQueryAsync();

            listBox4.Items.Add("СТАЛО:");
            listBox2.Items.Add("БЫЛО:");
            MainFrame.Content = new Page1();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (EXCEPTION.Visibility == Visibility.Visible)
                    EXCEPTION.Visibility = Visibility.Hidden;

                if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrWhiteSpace(textBox1.Text) &&
                    !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrWhiteSpace(textBox2.Text) &&
                    !string.IsNullOrEmpty(textBox3.Text) && !string.IsNullOrWhiteSpace(textBox3.Text) &&
                    !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrWhiteSpace(textBox4.Text) &&
                    !string.IsNullOrEmpty(textBox5.Text) && !string.IsNullOrWhiteSpace(textBox5.Text) &&
                    !string.IsNullOrEmpty(textBox6.Text) && !string.IsNullOrWhiteSpace(textBox6.Text) &&
                    !string.IsNullOrEmpty(textBox7.Text) && !string.IsNullOrWhiteSpace(textBox7.Text) &&
                    !string.IsNullOrEmpty(textBoxid.Text) && !string.IsNullOrWhiteSpace(textBoxid.Text))
                {
                    SqlCommand command = new SqlCommand("INSERT INTO InformationSecurityThreats1 (ИдентификаторУБИ, НаименованиеУБИ, Описание, ИсточникУгрозы, ОбъектВоздействияУгрозы, " +
                    "НарушениеКонфиденциальности, НарушениеЦелостности, НарушениеДоступности)VALUES(@ИдентификаторУБИ, @НаименованиеУБИ, @Описание, @ИсточникУгрозы," +
                    "@ОбъектВоздействияУгрозы, @НарушениеКонфиденциальности, @НарушениеЦелостности, @НарушениеДоступности)", sqlConnection);

                    command.Parameters.AddWithValue("ИдентификаторУБИ", textBoxid.Text);
                    command.Parameters.AddWithValue("НаименованиеУБИ", textBox1.Text);
                    command.Parameters.AddWithValue("Описание", textBox7.Text);
                    command.Parameters.AddWithValue("ИсточникУгрозы", textBox2.Text);
                    command.Parameters.AddWithValue("ОбъектВоздействияУгрозы", textBox3.Text);
                    command.Parameters.AddWithValue("НарушениеКонфиденциальности", textBox4.Text);
                    command.Parameters.AddWithValue("НарушениеЦелостности", textBox5.Text);
                    command.Parameters.AddWithValue("НарушениеДоступности", textBox6.Text);

                    await command.ExecuteNonQueryAsync();

                    MessageBox.Show("Успешно!".ToString(), "Успешно!".ToString(), MessageBoxButton.OK);
                    textBoxid.Clear();
                    textBox1.Clear();
                    textBox2.Clear();
                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox7.Clear();
                }
                else
                {
                    EXCEPTION.Visibility = Visibility.Visible;
                    EXCEPTION.Content = "Какое-то поле не заполненно!";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void Button_Click_1(object sender, RoutedEventArgs e)
        {
            SqlDataReader sqlReader = null;

            try
            {
                sqlReader = null;
                List<object> MyList = new List<object>();
                SqlCommand command = new SqlCommand("SELECT * FROM InformationSecurityThreats1", sqlConnection);
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    MyList.Add(new
                    {
                        ID = Convert.ToString(sqlReader["id"]),
                        IdYBI = Convert.ToString(sqlReader["ИдентификаторУБИ"]),
                        Name = Convert.ToString(sqlReader["НаименованиеУБИ"]),
                        Info = Convert.ToString(sqlReader["Описание"]),
                        Treat = Convert.ToString(sqlReader["ИсточникУгрозы"]),
                        ObjectTreat = Convert.ToString(sqlReader["ОбъектВоздействияУгрозы"]),
                        Confing = Convert.ToString(sqlReader["НарушениеКонфиденциальности"]),
                        Cel = Convert.ToString(sqlReader["НарушениеЦелостности"]),
                        Dost = Convert.ToString(sqlReader["НарушениеДоступности"])
                    });

                }
                listBox3.ItemsSource = MyList;

                MessageBox.Show(update_sum.ToString(), "Количество обновленных записей.", MessageBoxButton.OK);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
            }

        }

        private async void Button_Click_2(object sender, RoutedEventArgs e)
        {
            SqlDataReader sqlReader = null;
            if (EXCEPTION1.Visibility == Visibility.Visible)
                EXCEPTION1.Visibility = Visibility.Hidden;

            if (!string.IsNullOrEmpty(textBoxID.Text) && !string.IsNullOrWhiteSpace(textBoxID.Text) &&
                !string.IsNullOrEmpty(textBox8.Text) && !string.IsNullOrWhiteSpace(textBox8.Text) &&
                !string.IsNullOrEmpty(textBox9.Text) && !string.IsNullOrWhiteSpace(textBox9.Text) &&
                !string.IsNullOrEmpty(textBox10.Text) && !string.IsNullOrWhiteSpace(textBox10.Text) &&
                !string.IsNullOrEmpty(textBox11.Text) && !string.IsNullOrWhiteSpace(textBox11.Text) &&
                !string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrWhiteSpace(textBox12.Text) &&
                !string.IsNullOrEmpty(textBox13.Text) && !string.IsNullOrWhiteSpace(textBox13.Text) &&
                !string.IsNullOrEmpty(textBox14.Text) && !string.IsNullOrWhiteSpace(textBox14.Text) &&
                !string.IsNullOrEmpty(textBox15.Text) && !string.IsNullOrWhiteSpace(textBox15.Text))
            {
                SqlCommand command = new SqlCommand("SELECT * FROM InformationSecurityThreats1", sqlConnection);
                try
                {
                    sqlReader = await command.ExecuteReaderAsync();
                    while (await sqlReader.ReadAsync())
                    {
                        if (Convert.ToString(sqlReader["id"]) == textBoxID.Text)
                        {
                            Before.Add(Convert.ToString(sqlReader["ИдентификаторУБИ"]));
                            Before.Add(Convert.ToString(sqlReader["НаименованиеУБИ"]));
                            Before.Add(Convert.ToString(sqlReader["Описание"]));
                            Before.Add(Convert.ToString(sqlReader["ИсточникУгрозы"]));
                            Before.Add(Convert.ToString(sqlReader["ОбъектВоздействияУгрозы"]));
                            Before.Add(Convert.ToString(sqlReader["НарушениеКонфиденциальности"]));
                            Before.Add(Convert.ToString(sqlReader["НарушениеЦелостности"]));
                            Before.Add(Convert.ToString(sqlReader["НарушениеДоступности"]));
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                }
                finally
                {
                    if (sqlReader != null)
                        sqlReader.Close();
                }

                try
                {
                    command = new SqlCommand("UPDATE InformationSecurityThreats1 SET [ИдентификаторУБИ]=@ИдентификаторУБИ, [НаименованиеУБИ]=@НаименованиеУБИ, [Описание]=@Описание, [ИсточникУгрозы]=@ИсточникУгрозы, " +
                            "[ОбъектВоздействияУгрозы]=@ОбъектВоздействияУгрозы, [НарушениеКонфиденциальности]=@НарушениеКонфиденциальности, " +
                            "[НарушениеЦелостности]=@НарушениеЦелостности, [НарушениеДоступности]=@НарушениеДоступности WHERE [id]=@id", sqlConnection);

                    command.Parameters.AddWithValue("id", textBoxID.Text);
                    command.Parameters.AddWithValue("ИдентификаторУБИ", textBox9.Text);
                    command.Parameters.AddWithValue("НаименованиеУБИ", textBox10.Text);
                    command.Parameters.AddWithValue("Описание", textBox8.Text);
                    command.Parameters.AddWithValue("ИсточникУгрозы", textBox11.Text);
                    command.Parameters.AddWithValue("ОбъектВоздействияУгрозы", textBox12.Text);
                    command.Parameters.AddWithValue("НарушениеКонфиденциальности", textBox13.Text);
                    command.Parameters.AddWithValue("НарушениеЦелостности", textBox14.Text);
                    command.Parameters.AddWithValue("НарушениеДоступности", textBox15.Text);

                    await command.ExecuteNonQueryAsync();

                    listBox4.Items.Add($"({textBox9.Text}) ({textBox10.Text}) ({textBox8.Text}) " +
                                    $"({textBox11.Text}) ({textBox12.Text}) ({textBox13.Text}) " +
                                    $"({textBox14.Text}) ({textBox15.Text})");

                    string str = "";
                    foreach (object item in Before)
                    {
                        str += $"({item}) ";
                    }
                    Before.Clear();

                    listBox2.Items.Add(str);

                    textBoxID.Clear();
                    textBox8.Clear();
                    textBox9.Clear();
                    textBox10.Clear();
                    textBox11.Clear();
                    textBox12.Clear();
                    textBox13.Clear();
                    textBox14.Clear();
                    textBox15.Clear();

                    MessageBox.Show("Успешно!", "Сведения об операции", MessageBoxButton.OK);
                    update_sum++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                EXCEPTION1.Visibility = Visibility.Visible;
                EXCEPTION1.Content = "Заполните поле!";
            }
        }

        private async void Button_Click_3(object sender, RoutedEventArgs e)
        {
            if (EXCEPTION2.Visibility == Visibility.Visible)
                EXCEPTION2.Visibility = Visibility.Hidden;

            if (EZ.Visibility == Visibility.Visible)
                EZ.Visibility = Visibility.Hidden;

            try
            {
                if (!string.IsNullOrEmpty(textBox16.Text) && !string.IsNullOrWhiteSpace(textBox16.Text))
                {
                    SqlCommand command = new SqlCommand("DELETE FROM InformationSecurityThreats1 WHERE [ИдентификаторУБИ]=@ИдентификаторУБИ", sqlConnection);
                    command.Parameters.AddWithValue("ИдентификаторУБИ", textBox16.Text);


                    await command.ExecuteNonQueryAsync();

                    MessageBox.Show("Успешно!", "Сведения об операции", MessageBoxButton.OK);
                    EZ.Visibility = Visibility.Visible;
                    textBox16.Clear();
                }
                else
                {
                    EXCEPTION2.Visibility = Visibility.Visible;
                    EXCEPTION2.Content = "Необходимо вписать Идентификатор УБИ!!!";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

     

        private async void Button_Click_4(object sender, RoutedEventArgs e)
        {
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\User\source\repos\LABA2\LABA2\Database2.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);
            await sqlConnection.OpenAsync();
            SqlCommand command = new SqlCommand("SELECT * FROM InformationSecurityThreats1", sqlConnection);
            SqlDataReader sqlReader = null;
            FileStream file = new FileStream(@"C:\Users\User\Desktop\DB\table.txt", FileMode.Append);
            StreamWriter stream = new StreamWriter(file);
            try
            {
                sqlReader = await command.ExecuteReaderAsync();
                while (await sqlReader.ReadAsync())
                {
                    stream.WriteLine($"({Convert.ToString(sqlReader["id"])}) ({Convert.ToString(sqlReader["ИдентификаторУБИ"])}) ({Convert.ToString(sqlReader["НаименованиеУБИ"])}) ({Convert.ToString(sqlReader["Описание"])}) " +
                                $"({Convert.ToString(sqlReader["ИсточникУгрозы"])}) ({Convert.ToString(sqlReader["ОбъектВоздействияУгрозы"])}) ({Convert.ToString(sqlReader["НарушениеКонфиденциальности"])}) " +
                                $"({Convert.ToString(sqlReader["НарушениеЦелостности"])}) ({Convert.ToString(sqlReader["НарушениеДоступности"])})");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                stream.Close();
                file.Close();
                if (sqlReader != null)
                    sqlReader.Close();
            }
        }

        private async void Button_Click_5(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkbook;
            Excel.Worksheet xlWorksheet;
            Excel.Range xlRange;

            int xlRow;
            string strFileName;
            OpenFileDialog openFD = new OpenFileDialog();

            try
            {
                openFD.Filter = "Excel office |*.xls; *xlsx";
                openFD.ShowDialog();
                strFileName = openFD.FileName;

                if (strFileName != null)
                {
                    SqlCommand command;
                    xlApp = new Excel.Application();

                    try
                    {
                        command = new SqlCommand("DELETE FROM InformationSecurityThreats1 WHERE id < 100000", sqlConnection);

                        await command.ExecuteNonQueryAsync();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    finally
                    {
                        xlWorkbook = xlApp.Workbooks.Open(strFileName);
                        xlWorksheet = xlWorkbook.Worksheets["Sheet"];
                        xlRange = xlWorksheet.UsedRange;
                        for (xlRow = 3; xlRow <= xlRange.Rows.Count; xlRow++)
                        {
                            command = new SqlCommand("INSERT INTO InformationSecurityThreats1 (ИдентификаторУБИ, НаименованиеУБИ, Описание, ИсточникУгрозы, ОбъектВоздействияУгрозы, " +
                           "НарушениеКонфиденциальности, НарушениеЦелостности, НарушениеДоступности)VALUES(@ИдентификаторУБИ, @НаименованиеУБИ, @Описание, @ИсточникУгрозы," +
                           "@ОбъектВоздействияУгрозы, @НарушениеКонфиденциальности, @НарушениеЦелостности, @НарушениеДоступности)", sqlConnection);
                            command.Parameters.AddWithValue("@ИдентификаторУБИ", xlRange.Cells[xlRow, 1].Text);
                            command.Parameters.AddWithValue("@НаименованиеУБИ", xlRange.Cells[xlRow, 2].Text);
                            command.Parameters.AddWithValue("@Описание", xlRange.Cells[xlRow, 3].Text);
                            command.Parameters.AddWithValue("@ИсточникУгрозы", xlRange.Cells[xlRow, 4].Text);
                            command.Parameters.AddWithValue("@ОбъектВоздействияУгрозы", xlRange.Cells[xlRow, 5].Text);
                            command.Parameters.AddWithValue("@НарушениеКонфиденциальности", xlRange.Cells[xlRow, 6].Text);
                            command.Parameters.AddWithValue("@НарушениеЦелостности", xlRange.Cells[xlRow, 7].Text);
                            command.Parameters.AddWithValue("@НарушениеДоступности", xlRange.Cells[xlRow, 8].Text);

                            await command.ExecuteNonQueryAsync();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}