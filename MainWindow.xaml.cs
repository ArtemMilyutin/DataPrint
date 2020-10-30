using System;
using System.Collections.Generic;
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
using System.Threading;
using System.Printing;
using MySql.Data.MySqlClient;
using System.Data;
using System.IO;

namespace DataPrint
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Thread thr;
        MySqlConnectionStringBuilder mysqlCSB;
        DataTable dt = new DataTable();
        DataRow dr;

        int bufID =10000;
        int bufSize = 0;
        int bufPages = 0;
        bool flagFilter = false;

        public MainWindow()
        {
            try
            {
                InitializeComponent();
                mysqlCSB = new MySqlConnectionStringBuilder();
                mysqlCSB.Server = "xxx.xxx.xxx.xxx";
                mysqlCSB.Database = "base_name";
                mysqlCSB.UserID = "admin";
                mysqlCSB.Password = "password";
                mysqlCSB.CharacterSet = "utf8";


                dt.Columns.Add(new DataColumn("Имя пользователя"));
                dt.Columns.Add(new DataColumn("Принтер"));
                dt.Columns.Add(new DataColumn("Наименование файла"));
                dt.Columns.Add(new DataColumn("Кол-во страниц"));
                dt.Columns.Add(new DataColumn("Размер файла (Кб)"));
                dt.Columns.Add(new DataColumn("Дата печати"));

            }
            catch
            {
                MessageBox.Show("Ошибка подключения к базе данных!");
                Application.Current.Shutdown();
            }
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            #region вывод информации в ДатаГрид
            RefreshData();
            #endregion
            thr = new Thread(PrinterReading);
            thr.Start();

        }
        

        private void PrinterReading()
        {
            while (true)
            {
                try
                {
                    #region PrintServer
                    
                    string printServerName = @"\\serverName";

                    PrintServer ps = string.IsNullOrEmpty(printServerName)
                        // for local printers
                        ? new PrintServer()
                        // for shared printers
                        : new PrintServer(printServerName);
                    ps.EventLog.ToString();
                    PrintQueueCollection pqs = ps.GetPrintQueues();
                    string queryString;
                    MySqlConnection con;

                    foreach (PrintQueue pq in pqs)
                    {
                        if (!pq.IsOffline)
                        {
                            try
                            {
                                if (pq.NumberOfJobs > 0)
                                {
                                    if (pq.GetPrintJobInfoCollection().Count() != 0)
                                    {
                                        foreach (PrintSystemJobInfo pic in pq.GetPrintJobInfoCollection())
                                        {
                                            if (pic.TimeJobSubmitted.ToLocalTime() >= DateTime.Today && pic.JobIdentifier < 1000)
                                            {
                                                /*byte[] bytes = Encoding.Default.GetBytes(pic.Name);
                                                string FileName = Encoding.UTF8.GetString(bytes);*/
                                                bool copy = false;
                                                queryString = "SELECT `id`,`datetime` FROM `print_info` WHERE `datetime`='"+ pic.TimeJobSubmitted.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + "' AND `filename`='"+pic.Name+"' ORDER BY `id` DESC LIMIT 1000";
                                                using (con = new MySqlConnection())
                                                {

                                                    con.ConnectionString = mysqlCSB.ConnectionString;
                                                    MySqlCommand com = new MySqlCommand(queryString, con);
                                                    con.Open();
                                                    com.ExecuteNonQuery();
                                                    MySqlDataReader rdr = com.ExecuteReader();

                                                    if (rdr.HasRows)
                                                    {
                                                        copy = true;                                                           
                                                    }
                                                    
                                                }
                                                if (copy==true)
                                                {
                                                    if (pic.JobSize > bufSize || pic.NumberOfPages > bufPages)
                                                    {

                                                        bufSize = pic.JobSize;
                                                        bufPages = pic.NumberOfPages;

                                                        queryString = @"UPDATE `print_info` t1, (SELECT MAX(`id`)as id FROM `print_info` WHERE `username`='"+ pic.Submitter + 
                                                            "' ) t2 SET t1.`pages`='" + pic.NumberOfPages + "',t1.`size`='" + pic.JobSize / 1024 + 
                                                            "' WHERE `username`='" + pic.Submitter + "' AND t1.`id`=t2.`id` ";
                                                            
                                                        using (con = new MySqlConnection())
                                                        {
                                                            con.ConnectionString = mysqlCSB.ConnectionString;
                                                            MySqlCommand com = new MySqlCommand(queryString, con);
                                                            con.Open();
                                                            com.ExecuteNonQuery();                                                           
                                                            con.Close();
                                                        }
                                                    }
                                                }                                                
                                                else
                                                {
                                                    
                                                    bufID = pic.JobIdentifier;
                                                    bufSize = pic.JobSize;
                                                    bufPages = pic.NumberOfPages;


                                                    queryString = @"INSERT INTO `print_info` (`username`,`printer`,`filename`,`pages`,`size`,`datetime`) VALUES ('" + pic.Submitter
                                                        + "','" + pq.Name + "','" + pic.Name + "','" + pic.NumberOfPages + "','" + pic.JobSize / 1024 + "','" + pic.TimeJobSubmitted.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") + "')";
                                                    using (con = new MySqlConnection())
                                                    {
                                                        con.ConnectionString = mysqlCSB.ConnectionString;
                                                        MySqlCommand com = new MySqlCommand(queryString, con);
                                                        con.Open();
                                                        com.ExecuteNonQuery();
                                                        con.Close();
                                                    }
                                                }
                                                string str = pic.JobIdentifier + ". " + pic.Submitter + "\t" + pq.Name + "\t" + pic.Name + "\t "
                                                        + pic.NumberOfPages + "\t " + pic.JobStatus + "\t " + pic.JobSize / 1024 + " Kb\t " + pic.TimeJobSubmitted.ToLocalTime().ToString() ;
                                                

                                                using (StreamWriter sw = new StreamWriter("log.txt", true))
                                                {
                                                    sw.WriteLine(str);
                                                }

                                                if (flagFilter==false)
                                                    RefreshData();
                                            }
                                        }                                        
                                    }
                                }
                                #region Информация о принтере (IP, Имя, домен...) 
                                
                                #endregion
                            }
                            catch {
                                //thr.Abort();
                            }
                        }
                        
                    }
                    #endregion
                }
                catch { }
                
            }
        }

        void RefreshData()
        {
            dt = new DataTable();
            dt.Columns.Add(new DataColumn("Имя пользователя"));
            dt.Columns.Add(new DataColumn("Принтер"));
            dt.Columns.Add(new DataColumn("Наименование файла"));
            dt.Columns.Add(new DataColumn("Кол-во страниц"));
            dt.Columns.Add(new DataColumn("Размер файла (Кб)"));
            dt.Columns.Add(new DataColumn("Дата печати"));

            MySqlConnection con;
            string queryString = "SELECT `username`,`printer`,`filename`,`pages`,`size`,`datetime` FROM `print_info` ORDER BY `id` DESC LIMIT 1000";
            using (con = new MySqlConnection())
            {

                con.ConnectionString = mysqlCSB.ConnectionString;
                MySqlCommand com = new MySqlCommand(queryString, con);
                con.Open();
                com.ExecuteNonQuery();
                MySqlDataReader rdr = com.ExecuteReader();


                while (rdr.Read())
                {

                    dr = dt.NewRow();

                    dr["Имя пользователя"] = rdr.GetString(0);
                    dr["Принтер"] = rdr.GetString(1);
                    dr["Наименование файла"] = rdr.GetString(2);
                    dr["Кол-во страниц"] = rdr.GetString(3);
                    dr["Размер файла (Кб)"] = rdr.GetString(4);
                    DateTime date = new DateTime(rdr.GetMySqlDateTime(5).Year, rdr.GetMySqlDateTime(5).Month, rdr.GetMySqlDateTime(5).Day,
                                              rdr.GetMySqlDateTime(5).Hour, rdr.GetMySqlDateTime(5).Minute, rdr.GetMySqlDateTime(5).Second);

                    dr["Дата печати"] = date.ToString("dd.MM.yyyy HH:mm:ss");


                    dt.Rows.Add(dr);
                }
                
                dGrid.Dispatcher.Invoke(new Action(()=> { dGrid.ItemsSource = dt.DefaultView;  }));


                con.Close();
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            thr.Abort();
        }

        private void Filter()
        {
            flagFilter = true;

            dt = new DataTable();
            dt.Columns.Add(new DataColumn("Имя пользователя"));
            dt.Columns.Add(new DataColumn("Принтер"));
            dt.Columns.Add(new DataColumn("Наименование файла"));
            dt.Columns.Add(new DataColumn("Кол-во страниц"));
            dt.Columns.Add(new DataColumn("Размер файла (Кб)"));
            dt.Columns.Add(new DataColumn("Дата печати"));

            string stringWhere = "";
            if (textBox1.Text == String.Empty && datePickerMin.Text == String.Empty && datePickerMax.Text == String.Empty)//не указаны все
                stringWhere = "";
            else if (textBox1.Text != String.Empty && datePickerMin.Text != String.Empty && datePickerMax.Text != String.Empty)//указаны все
                stringWhere = " WHERE `username`='" + textBox1.Text + "' AND `datetime`>='" + datePickerMin.SelectedDate.Value.ToString("yyyy-MM-dd") + "' AND `datetime`<='" + datePickerMax.SelectedDate.Value.ToString("yyyy-MM-dd") + "' ";
            else if (textBox1.Text == String.Empty && datePickerMin.Text != String.Empty && datePickerMax.Text != String.Empty)//указаны min и max
                stringWhere = " WHERE `datetime`>='" + datePickerMin.SelectedDate.Value.ToString("yyyy-MM-dd") + "' AND `datetime`<='" + datePickerMax.SelectedDate.Value.ToString("yyyy-MM-dd") + "' ";
            else if (textBox1.Text != String.Empty && datePickerMin.Text == String.Empty && datePickerMax.Text != String.Empty)//указаны имя и max
                stringWhere = " WHERE `username`='" + textBox1.Text + "' AND `datetime`<='" + datePickerMax.SelectedDate.Value.ToString("yyyy-MM-dd") + "' ";
            else if (textBox1.Text != String.Empty && datePickerMin.Text != String.Empty && datePickerMax.Text == String.Empty)//указаны min и имя
                stringWhere = " WHERE `username`='" + textBox1.Text + "' AND `datetime`>='" + datePickerMin.SelectedDate.Value.ToString("yyyy-MM-dd") + "' ";
            else if (textBox1.Text != String.Empty && datePickerMin.Text == String.Empty && datePickerMax.Text == String.Empty)//указано только имя
                stringWhere = " WHERE `username`='" + textBox1.Text + "' ";
            else if (textBox1.Text == String.Empty && datePickerMin.Text != String.Empty && datePickerMax.Text == String.Empty)//указан только min 
                stringWhere = " WHERE `datetime`>='" + datePickerMin.SelectedDate.Value.ToString("yyyy-MM-dd") + "' ";
            else if (textBox1.Text == String.Empty && datePickerMin.Text == String.Empty && datePickerMax.Text != String.Empty)//указан только max 
                stringWhere = " WHERE `datetime`<='" + datePickerMax.SelectedDate.Value.ToString("yyyy-MM-dd") + "' ";

            if (stringWhere != "")
            {
                try
                {
                    MySqlConnection con;
                    string queryString = "SELECT `username`,`printer`,`filename`,`pages`,`size`,`datetime` FROM `print_info` " + stringWhere + " ORDER BY `id` DESC ";
                    using (con = new MySqlConnection())
                    {

                        con.ConnectionString = mysqlCSB.ConnectionString;
                        MySqlCommand com = new MySqlCommand(queryString, con);
                        con.Open();
                        com.ExecuteNonQuery();
                        MySqlDataReader rdr = com.ExecuteReader();

                        while (rdr.Read())
                        {
                            dr = dt.NewRow();

                            dr["Имя пользователя"] = rdr.GetString(0);
                            dr["Принтер"] = rdr.GetString(1);
                            dr["Наименование файла"] = rdr.GetString(2);
                            dr["Кол-во страниц"] = rdr.GetString(3);
                            dr["Размер файла (Кб)"] = rdr.GetString(4);
                            DateTime date = new DateTime(rdr.GetMySqlDateTime(5).Year, rdr.GetMySqlDateTime(5).Month, rdr.GetMySqlDateTime(5).Day,
                                                      rdr.GetMySqlDateTime(5).Hour, rdr.GetMySqlDateTime(5).Minute, rdr.GetMySqlDateTime(5).Second);
                            dr["Дата печати"] = date.ToString("dd.MM.yyyy HH:mm:ss");

                            dt.Rows.Add(dr);
                        }
                        con.Close();
                    }
                    dGrid.Dispatcher.Invoke(new Action(() => { dGrid.ItemsSource = dt.DefaultView; }));
                }
                catch { MessageBox.Show("Ошибка соединения с сервером."); }
            }
            else
            {
                flagFilter = false;
                datePickerMin.Text = "";
                datePickerMax.Text = "";
                textBox1.Text = "";

                RefreshData();
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)//сброс фильтров
        {
            //label1.Dispatcher.Invoke(new Action(() => { label1.Content=""; }));

            flagFilter = false;
            datePickerMin.Text = "";
            datePickerMax.Text = "";
            textBox1.Text = "";

            RefreshData();
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)//фильтрация
        {
            Filter();
        }

        private void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            PrintDialog printDialog = new PrintDialog();
            if (printDialog.ShowDialog() == true)
            {
                FlowDocument fd = new FlowDocument();

                Paragraph p = new Paragraph(new Run("Отчет системы мониторинга печати"));
                p.FontStyle = dGrid.FontStyle;
                p.FontFamily = dGrid.FontFamily;
                p.FontSize = 18;
                fd.Blocks.Add(p);

                Table table = new Table();
                TableRowGroup tableRowGroup = new TableRowGroup();
                TableRow r = new TableRow();
                fd.PageWidth = printDialog.PrintableAreaWidth;
                fd.PageHeight = printDialog.PrintableAreaHeight;
                fd.BringIntoView();

                fd.TextAlignment = TextAlignment.Center;
                fd.ColumnWidth = 450;
                table.CellSpacing = 0;

                var headerList = dGrid.Columns.Select(x => x.Header.ToString()).ToList();
                List<dynamic> bindList = new List<dynamic>();

                for (int j = 0; j < headerList.Count; j++)
                {
                    r.Cells.Add(new TableCell(new Paragraph(new Run(headerList[j]))));
                    if (j == 3 || j == 4) r.Cells[j].ColumnSpan = 2;
                    else if(j==2) r.Cells[j].ColumnSpan = 7;
                    else r.Cells[j].ColumnSpan = 4;

                    r.Cells[j].Padding = new Thickness(4);



                    r.Cells[j].BorderBrush = Brushes.Black;
                    r.Cells[j].FontWeight = FontWeights.Bold;
                    r.Cells[j].Background = Brushes.DarkGray;
                    r.Cells[j].Foreground = Brushes.White;
                    r.Cells[j].BorderThickness = new Thickness(1, 1, 1, 1);
                    var binding = (dGrid.Columns[j] as DataGridBoundColumn).Binding as Binding;

                    bindList.Add(binding.Path.Path);
                }
                tableRowGroup.Rows.Add(r);
                table.RowGroups.Add(tableRowGroup);
                for (int i = 0; i < dGrid.Items.Count-1; i++)
                {

                    DataRowView row = (DataRowView)dGrid.Items.GetItemAt(i);
                    
                    table.BorderBrush = Brushes.Gray;
                    table.BorderThickness = new Thickness(1, 1, 0, 0);
                    table.FontStyle = dGrid.FontStyle;
                    table.FontFamily = dGrid.FontFamily;
                    table.FontSize = 13;
                    tableRowGroup = new TableRowGroup();
                    r = new TableRow();
                    for (int j = 0; j < row.Row.ItemArray.Count(); j++)
                    {
                        r.Cells.Add(new TableCell(new Paragraph(new Run(row.Row.ItemArray[j].ToString()))));
                        
                        if (j == 3 || j == 4) r.Cells[j].ColumnSpan = 2;
                        else if (j == 2) r.Cells[j].ColumnSpan = 7;
                        else r.Cells[j].ColumnSpan = 4;
                        r.Cells[j].Padding = new Thickness(4);

                        r.Cells[j].BorderBrush = Brushes.DarkGray;
                        r.Cells[j].BorderThickness = new Thickness(0, 0, 1, 1);
                    }

                    tableRowGroup.Rows.Add(r);
                    table.RowGroups.Add(tableRowGroup);

                }
                fd.Blocks.Add(table);

                printDialog.PrintDocument(((IDocumentPaginatorSource)fd).DocumentPaginator, "Отчет системы мониторинга печати");

            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                if (textBox1.IsFocused)
                {
                    //label1.Dispatcher.Invoke(new Action(() => { label1.Content += "\nEnter"; }));
                    Filter();
                }
                else {
                    //label1.Dispatcher.Invoke(new Action(() => { label1.Content += "\nNotEnter"; }));
                }
            }
        }
    }
}
