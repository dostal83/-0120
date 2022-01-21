using System;
using System.Collections.Generic;
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
    public partial class Page3 : Page
    {
        SqlConnection sqlConnection;
        public Page3()
        {
            InitializeComponent();
        }

        private async void Page3_Loaded(object sender, RoutedEventArgs e)
        {
            string connectionString = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\QQQ\Desktop\LABA2\LABA2\Database1.mdf;Integrated Security=True";
            sqlConnection = new SqlConnection(connectionString);

            await sqlConnection.OpenAsync();
            SqlDataReader sqlReader = null;
            SqlCommand command = new SqlCommand("SELECT COUNT(*) FROM [InformationSecurityThreats1]", sqlConnection);
            try
            {
                listBoxCount3.Items.Clear();
                command = new SqlCommand("SELECT COUNT(*) FROM [InformationSecurityThreats1]", sqlConnection);
                int count = 0;
                count = (int)(await command.ExecuteScalarAsync());
                listBoxCount3.Items.Add($"Общее число записей: {count}");
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
                listBoxPage3.Items.Clear();
                listBoxPage3.Items.Add($"Идентификатор УБИ   Наименование УБИ");
                command = new SqlCommand("SELECT ИдентификаторУБИ, НаименованиеУБИ FROM [InformationSecurityThreats1]", sqlConnection);
                sqlReader = await command.ExecuteReaderAsync();
                int i = 200;
                while (i > 0 && await sqlReader.ReadAsync())
                {
                    i--;
                }

                i = 100;
                while (i > 0 && await sqlReader.ReadAsync())
                {
                    if (Convert.ToInt32(sqlReader["ИдентификаторУБИ"]) <= 9)
                    {
                        listBoxPage3.Items.Add($"УБИ.00{Convert.ToString(sqlReader["ИдентификаторУБИ"])}   {Convert.ToString(sqlReader["НаименованиеУБИ"])}");
                    }
                    else if (Convert.ToInt32(sqlReader["ИдентификаторУБИ"]) >= 10 && Convert.ToInt32(sqlReader["ИдентификаторУБИ"]) <= 99)
                    {
                        listBoxPage3.Items.Add($"УБИ.0{Convert.ToString(sqlReader["ИдентификаторУБИ"])}   {Convert.ToString(sqlReader["НаименованиеУБИ"])}");
                    }
                    else
                    {
                        listBoxPage3.Items.Add($"УБИ.{Convert.ToString(sqlReader["ИдентификаторУБИ"])}   {Convert.ToString(sqlReader["НаименованиеУБИ"])}");
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
        }

        private void But3_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page4());
        }

        private void But4_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Page2());
        }
    }
}
