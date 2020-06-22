using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
using excel = Microsoft.Office.Interop.Excel;

namespace FirstWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            excel.Application ap = new excel.Application();

            excel.Workbook wb = ap.Workbooks.Add(Missing.Value);

            excel.Worksheet ws = wb.Worksheets[1];
            int row = 0;

            ws.Cells[1, 1] = "Driver Name";
            ws.Cells[1, 2] = "Driver Family";
            ws.Cells[1, 3] = "Driver Phone";
            


            for (int i = 1; i < 3; i++)
            {
                row++;
                //ws.Range["a" + row].Value = i.ToString();
                ws.Cells[row+1, row+1] ="Text " + i.ToString();
            }
            ap.Visible = true;

        }
    }
}
