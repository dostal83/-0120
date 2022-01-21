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
    /// Логика взаимодействия для PageDT2.xaml
    /// </summary>
    public partial class PageDT2 : Page
    {
        public PageDT2()
        {
            InitializeComponent();
        }
        private async void PageDT_Loaded(object sender, RoutedEventArgs e)
        {
            SqlConnection sqlConnection = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\QQQ\Desktop\LABA2\LABA2\Database1.mdf;Integrated Security=True");
            await sqlConnection.OpenAsync();
            SqlCommand command = new SqlCommand("SELECT ИдентификаторУБИ, НаименованиеУБИ FROM [InformationSecurityThreats1] WHERE [ИдентификаторУБИ]<201 AND [ИдентификаторУБИ]>100", sqlConnection);
            await command.ExecuteNonQueryAsync();
            try
            {
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(command);
                DataTable dt = new DataTable("InformationSecurityThreats");
                sqlDataAdapter.Fill(dt);
                mygrid.ItemsSource = dt.DefaultView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), ex.Source.ToString(), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            try
            {
                listBoxCount.Items.Clear();
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
                sqlConnection.Close();
            }
        }
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageDT());
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new PageDT3());
        }
    }
}
