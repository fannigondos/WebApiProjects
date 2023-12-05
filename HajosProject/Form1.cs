using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.EntityFrameworkCore.Migrations.Operations.Builders;

namespace HajosProject
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        public Form1()
        {
            InitializeComponent();
            CreateExcel();

        }

        private void CreateExcel()
        {
            try
            {
                xlApp = new Excel.Application();
                xlWB = xlApp.Workbooks.Add(Missing.Value);
                xlSheet = xlWB.ActiveSheet;

                CreateTable();

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        private void CreateTable()
        {
            string[] fejlécek = new string[] {
                "Kérdés",
                "1. válasz",
                "2. válaszl",
                "3. válasz",
                "Helyes válasz",
                "kép"};

            for (int i = 0; i < fejlécek.Length; i++)
            {
                xlSheet.Cells[1, i + 1] = fejlécek[i];
            }

            Models.HajosContext hajosContext = new Models.HajosContext();
            var mindenkérdés = hajosContext.Questions.ToList();

            object[,] adatTömb = new object[mindenkérdés.Count(), fejlécek.Count()];

            for (int i = 1; i < mindenkérdés.Count(); i++)
            {
                adatTömb[i, 0] = mindenkérdés[i].Question1;
                adatTömb[i, 1] = mindenkérdés[i].Answer1;
                adatTömb[i, 2] = mindenkérdés[i].Answer2;
                adatTömb[i, 3] = mindenkérdés[i].Answer3;
                adatTömb[i, 4] = mindenkérdés[i].CorrectAnswer;
                adatTömb[i, 5] = mindenkérdés[i].Image;
            }

            int sorokSzáma = adatTömb.GetLength(0);
            int oszlopokSzáma = adatTömb.GetLength(1);

            Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSzáma, oszlopokSzáma);
            adatRange.Value2 = adatTömb;

            adatRange.Columns.AutoFit();

            Excel.Range fejlécRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejlécRange.Font.Bold = true;
            fejlécRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejlécRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejlécRange.EntireColumn.AutoFit();
            fejlécRange.RowHeight = 40;
            fejlécRange.Interior.Color = Color.Fuchsia;
            fejlécRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range táblaRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(lastRowID, 6);
            táblaRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range elsõoszlopRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(lastRowID, 1);
            elsõoszlopRange.Font.Bold = true;
            elsõoszlopRange.Interior.Color = Color.Yellow;

            Excel.Range utolsóoszlopRange = xlSheet.get_Range("F1", Type.Missing).get_Resize(lastRowID, 6);
            utolsóoszlopRange.Font.Bold = true;
            utolsóoszlopRange.Interior.Color = Color.LightGreen;
        } 
    }
}