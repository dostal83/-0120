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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LABA2
{
    /// <summary>
    /// Логика взаимодействия для Page1.xaml
    /// </summary>
    public partial class Page1 : Page
    {
        SqlConnection sqlConnection;
        public Page1()
        {
            InitializeComponent();
        }

        private async void Page1_Loaded(object sender, RoutedEventArgs e)
        {
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\QQQ\Desktop\LABA2\LABA2\Database1.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT ИдентификаторУБИ, НаименованиеУБИ FROM [InformationSecurityThreats1]", sqlConnection);
            try
            {
                listBoxPage.Items.Clear();
                listBoxPage.Items.Add($"Идентификатор УБИ   Наименование УБИ");
                sqlReader = await command.ExecuteReaderAsync();
                int i = 100;
                while (i > 0 && await sqlReader.ReadAsync())
                {
                    if (Convert.ToInt32(sqlReader["ИдентификаторУБИ"]) <= 9)
                    {
                        listBoxPage.Items.Add($"УБИ.00{Convert.ToString(sqlReader["ИдентификаторУБИ"])}   {Convert.ToString(sqlReader["НаименованиеУБИ"])}");
                    }
                    else if (Convert.ToInt32(sqlReader["ИдентификаторУБИ"]) >= 10 && Convert.ToInt32(sqlReader["ИдентификаторУБИ"]) <= 99)
                    {
                        listBoxPage.Items.Add($"УБИ.0{Convert.ToString(sqlReader["ИдентификаторУБИ"])}   {Convert.ToString(sqlReader["НаименованиеУБИ"])}");
                    }
                    else
                    {
                        listBoxPage.Items.Add($"УБИ.{Convert.ToString(sqlReader["ИдентификаторУБИ"])}   {Convert.ToString(sqlReader["НаименованиеУБИ"])}");
                    }

                    i--;
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
                listBoxCount.Items.Clear();
                sqlReader = null;
                command = new SqlCommand("SELECT COUNT(*) FROM [InformationSecurityThreats1]", sqlConnection);
                int count = 0;
                count = (int)(await command.ExecuteScalarAsync());

                listBoxCount.Items.Add($"Общее число записей: {count}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (sqlReader != null)
                    sqlReader.Close();
                sqlConnection.Close();
            }
        }

        private void But_Click_Page_Next(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page2());
        }

        private void dtLEave_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageDT());
        }
    }
}
