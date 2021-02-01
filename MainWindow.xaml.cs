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
using System.Diagnostics;
//using Excel = Microsoft.Office.Interop.Excel;

namespace DataTable_Intima_
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {        
        //string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 8.0;HDR=Yes";
        
        private static ObservableCollection<DatTable> dataCollection = new ObservableCollection<DatTable>();
        private string path = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        
        
        public MainWindow()
        {
            InitializeComponent();
            readFile(path);
            saveDataInFile(path);
        }

        private void readFile(string path)
        {
            path = System.IO.Path.Combine(path, "Test.xlsx");
            var wb = new XLWorkbook(path);
            var ws = wb.Worksheet(1);
            var rows = ws.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                dataCollection.Add(new DatTable(row.Cell(1).Value.ToString(), row.Cell(2).Value.ToString(), 
                                     row.Cell(3).Value.ToString(), row.Cell(4).Value.ToString()));//fix crutch
            }
            var sortData = dataCollection.OrderBy(u => u.Value);
            table.ItemsSource = sortData;
            DataTable dt = new DataTable();
            
        }

        private void saveDataInFile(string filePath)
        {
            filePath = System.IO.Path.Combine(path, "Test.xlsx");
            var wb = new XLWorkbook(filePath);
            wb.Worksheet(2).Delete();
            var ws = wb.Worksheets.Add(2); //var ws = wb.AddWorksheet("2");
            ws.Cell("A1").Value = table.Items;
            ws.Columns().AdjustToContents();
            wb.SaveAs(filePath);
        }

    }
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


//почти рабочая тема
//table.SelectAllCells();
//table.ClipboardCopyMode = DataGridClipboardCopyMode.IncludeHeader;

//String resultat = (string)Clipboard.GetData(DataFormats.CommaSeparatedValue);

//String result = (string)Clipboard.GetData(DataFormats.Text);

//table.UnselectAllCells();
//Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
//dlg.FileName = "Export1";
//dlg.DefaultExt = ".text";
//dlg.Filter = "Excel documents (.xlsx)|*.xlsx";

//Nullable<bool> result1 = dlg.ShowDialog();
//if (result1 == true)
//{

//    string filename = dlg.FileName;

//    System.IO.StreamWriter file = new System.IO.StreamWriter(filename, false);
//    file.WriteLine(result);
//    file.Close();

//    MessageBox.Show("Экспорт данных успешно завершен");
//}


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