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
            string[] fejl�cek = new string[] {
                "K�rd�s",
                "1. v�lasz",
                "2. v�laszl",
                "3. v�lasz",
                "Helyes v�lasz",
                "k�p"};

            for (int i = 0; i < fejl�cek.Length; i++)
            {
                xlSheet.Cells[1, i + 1] = fejl�cek[i];
            }

            Models.HajosContext hajosContext = new Models.HajosContext();
            var mindenk�rd�s = hajosContext.Questions.ToList();

            object[,] adatT�mb = new object[mindenk�rd�s.Count(), fejl�cek.Count()];

            for (int i = 1; i < mindenk�rd�s.Count(); i++)
            {
                adatT�mb[i, 0] = mindenk�rd�s[i].Question1;
                adatT�mb[i, 1] = mindenk�rd�s[i].Answer1;
                adatT�mb[i, 2] = mindenk�rd�s[i].Answer2;
                adatT�mb[i, 3] = mindenk�rd�s[i].Answer3;
                adatT�mb[i, 4] = mindenk�rd�s[i].CorrectAnswer;
                adatT�mb[i, 5] = mindenk�rd�s[i].Image;
            }

            int sorokSz�ma = adatT�mb.GetLength(0);
            int oszlopokSz�ma = adatT�mb.GetLength(1);

            Excel.Range adatRange = xlSheet.get_Range("A2", Type.Missing).get_Resize(sorokSz�ma, oszlopokSz�ma);
            adatRange.Value2 = adatT�mb;

            adatRange.Columns.AutoFit();

            Excel.Range fejl�cRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(1, 6);
            fejl�cRange.Font.Bold = true;
            fejl�cRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            fejl�cRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            fejl�cRange.EntireColumn.AutoFit();
            fejl�cRange.RowHeight = 40;
            fejl�cRange.Interior.Color = Color.Fuchsia;
            fejl�cRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range t�blaRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(lastRowID, 6);
            t�blaRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            Excel.Range els�oszlopRange = xlSheet.get_Range("A1", Type.Missing).get_Resize(lastRowID, 1);
            els�oszlopRange.Font.Bold = true;
            els�oszlopRange.Interior.Color = Color.Yellow;

            Excel.Range utols�oszlopRange = xlSheet.get_Range("F1", Type.Missing).get_Resize(lastRowID, 6);
            utols�oszlopRange.Font.Bold = true;
            utols�oszlopRange.Interior.Color = Color.LightGreen;
        } 
    }
}