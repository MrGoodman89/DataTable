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
using System.Windows.Threading;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Data.OleDb;
using ClosedXML.Excel; 

//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
//using Microsoft.Office.Interop.Excel;

namespace DataTable_Intima_
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        //private static Collection<DatTable> dataCollection = new Collection<DatTable>();
        //string pathToFile = @"C:\Users\lifec\Desktop\Test.xlsx";
        //string[,] list;
        //List<List<string>> spisok;

        ////string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\\Users\\lifec\\Desktop\\Test.xlsx;Extended Properties=Excel 8.0;HDR=Yes";
        ////string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source = C:\Users\lifec\Desktop\Test.xlsx;Extended Properties = 'Excel 8.0;HDR=Yes;IMEX=1'";
        //static string extending = "Excel 8.0;HDR=Yes;IMEX=1";
        //string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\lifec\\Desktop\\Test.xlsx;Extended Properties=" + extending +"";
        //DataGrid dt = new DataGrid();
        public MainWindow()
        {
            InitializeComponent();
            //readFromExcel();
            //table.ItemsSource = dataCollection;
            //table.ItemsSource = dataCollection;
            //OpenConnection(connectionString);
            //table.ItemsSource = dataCollection;


        }

        //private void readFromExcel()
        //{
        //    Microsoft.Office.Interop.Excel.Application excle_app = new Microsoft.Office.Interop.Excel.Application();
        //    excle_app.Visible = true;
        //    Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = excle_app.Workbooks.Open(pathToFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        //                                                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    //Выбираем таблицу(лист).
        //    Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
        //    var lastCell = ObjWorkSheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell);
        //    list = new string[lastCell.Row, lastCell.Column];
        //    spisok = new List<List<string>>();
        //    //string[,] list = new string[lastCell.Column, lastCell.Row];
        //    dataCollection.Clear();
        //    for (int i = 0; i < (int)lastCell.Row; i++)
        //    {
        //        for (int j = 0; j < (int)lastCell.Column; j++)
        //        {
        //            list[i,j] = ObjWorkSheet.Cells[i + 1,j + 1].Text.ToString();
                    
        //            //list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString();//считал текст в строку
        //            //if (j == (int)lastCell.Column)
        //            //{
        //            //    //    spisok.Add(ObjWorkSheet.Cells[i + 1, j + 1].Text.ToString());
        //            //    for (int z = 0; z < list.Length; i++)
        //            //    {
        //            //        for (int x = 0; x < list.Length; x++)
        //            //        {
        //            //            dataCollection.Add(new DatTable(list[z, x].ToString(), list[z, x].ToString(), list[z, x].ToString(), list[z, x].ToString()));   
        //            //        }
        //            //    }
        //            //}
        //        }
        //    }
        //    //sortList(ref spisok);
      
        //    table.ItemsSource = spisok;
        //    table.Items.Refresh();


        //    ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
        //    excle_app.Quit(); // вышел из Excel
        //    //GC.Collect(); // убрал за собой

        //    //table.ItemsSource = list;

        //    //Microsoft.Office.Interop.Excel.Range usedColumn = ObjWorkSheet.UsedRange.Columns[numRow];
        //    //System.Array myvalues = (System.Array)usedColumn.Cells.Value2;
        //    //string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

        //    // Выходим из программы Excel.
        //    //excle_app.Quit();

        //}

        //public void sortList(ref List<List<string>> listok)
        //{
        //    listok.Sort(new DataComparer());
        //    foreach (var el1 in listok)
        //    {

        //    }
        //}
        public void addListInGrid()
        {

        }
        //private DataTable RequestProcessing(string QueryString)
        //{
        //    DataTable datatable = new DataTable();

        //    using (OleDbConnection connection = new OleDbConnection(connectionString))
        //    {
        //        DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
        //        DataRow schemaRow = schemaTable.Rows[0];

        //        return datatable;
        //        //SqlCommand SqlCommand = new SqlCommand(QueryString, SqlConnection);
        //        //SqlDataAdapter adp = new SqlDataAdapter(SqlCommand);
        //        //adp.Fill(datatable);
        //        //return datatable;
        //    }
        //}
        //private void readFromExcel()
        //{
        //    string sheet = schemaRow["TABLE_NAME"].ToString();
        //    DataTable dataTable = RequestProcessing("SELECT * FROM [" + sheet + "]");
        //    for (int i = 0; i < dataTable.Rows.Count; i++)
        //    {
        //        dataCollection.Add(new DatTable(dataTable.Rows[][], ));
        //    }

        //}


        //static void OpenConnection(string connectionString)
        //{
        //    using (OleDbConnection connection = new OleDbConnection(connectionString))
        //    {
        //        try
        //        {
        //            connection.Open();
        //            //Запрашиваем таблицы
        //            DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
        //            DataRow schemaRow = schemaTable.Rows[0];
        //            //Получаеи имя таблицы
        //            string sheet = schemaRow["TABLE_NAME"].ToString();
        //            //Объявляем команду
        //            OleDbCommand com = connection.CreateCommand();
        //            //Создаем SQL запрос
        //            com.CommandText = "SELECT * FROM [" + sheet + "]";
        //            //Выполняем SQL запрос
        //            OleDbDataReader reader = com.ExecuteReader();
        //            //Записываем результат в DataTable
        //            DataTable dTable = new DataTable();
        //            dTable.Load(reader);
        //            //Выводим DataTable в таблицу на форму (если нужно)
        //            for (int i = 0; i < dTable.Rows.Count; i++)
        //            {
        //                for (int j = 0; j < dTable.Columns.Count; i++)
        //                {
        //                    dataCollection.Add(new DatTable(dTable.Rows[i][j].ToString(), dTable.Rows[i][j].ToString(), dTable.Rows[i][j].ToString(), dTable.Rows[i][j].ToString()));
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex.Message);
        //        }
        //        // The connection is automatically closed when the
        //        // code exits the using block.
        //    }
        //}


        //static void ReadExcelFile(string fileName)
        //{
        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
        //    {
        //        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        //        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

        //        OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
        //        string text;
        //        while (reader.Read())
        //        {
        //            if (reader.ElementType == typeof(CellValue))
        //            {
        //                text = reader.GetText();
        //            }
        //        }
        //    }

        //}

        private void table_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }


    }
}


//ПОДКЛЮЧЕНИЕ ЧЕРЕЗ ОЛЕДБ (ВИДОС НА ЮТУБЕ)

//DataSet ds = new DataSet();
//OleDbDataAdapter adapter = new OleDbDataAdapter();

//private void getData_Click(object sender, RoutedEventArgs e)
//{
//    //string path = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
//    //path = System.IO.Path.Combine(path, "Test.xlsx");
//    //string connectingString = @"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" + path + @";Extended Properties=""Excel 12.0 Macro;HDR=Yes;ImpoertMixedTypes=Text;TypeGuessRows=0""";

//    //OleDbConnection conn = new OleDbConnection(connectingString);
//    //string strCmd = "select * from [Sheets1$A2:D10]";
//    //OleDbCommand cmd = new OleDbCommand(strCmd, conn);

//    //try
//    //{
//    //    conn.Open();
//    //    ds.Clear();
//    //    adapter.SelectCommand = cmd;
//    //    adapter.Fill(ds);
//    //    table.ItemsSource = (System.Collections.IEnumerable)ds.Tables[0];
//    //}
//    //catch(Exception ex)
//    //{
//    //    Console.WriteLine(ex.Message);
//    //}
//    //finally
//    //{
//    //    conn.Close();
//    //}
//}