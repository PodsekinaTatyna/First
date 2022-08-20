using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace CritPearsona
{

    public abstract class Data
    {
        public Data() { }

        public double[] GetData(double[,] masData2)
        {
            double[] masData1 = new double[masData2.GetLength(0)];
            for (int i = 0; i < masData1.Length; i++)
            {
                masData1[i] = masData2[i, 0];
            }
            return masData1;
        }

        public abstract void SortingData();
    }


    class WorkingWithData : Data
    {
        public double[] masData;
        public WorkingWithData(double[] masData)
        {
            this.masData = masData;
        }

        public WorkingWithData() { }

        public override void SortingData()
        {
            for (int i = 0; i < masData.Length; i++)
            {
                for (int j = 0; j < masData.Length - 1; j++)
                {
                    if (masData[j] > masData[j + 1])
                    {
                        double max = masData[j];
                        masData[j] = masData[j + 1];
                        masData[j + 1] = max;
                    }
                }
            }

        }

    }


    class RequiredValues : ICalculations
    {
        protected double[] masData;
        protected double N;
        protected double k;
        protected double C;
        protected double S;
        protected double scaleNumber;
        protected double averageX;

        public RequiredValues() { }

        public void Colculations()
        {
            N = masData.Length;
            for (int i = 0; i < masData.Length; i++)
            {
                averageX += masData[i];
            }
            averageX = averageX / masData.Length;
            k = Math.Ceiling(1 + 3.32 * Math.Log10(N));
            C = Math.Ceiling((masData[masData.Length - 1] - masData[0]) / k);
            for (int i = 0; i < masData.Length; i++)
            {
                S += Math.Pow(masData[i] - averageX, 2);
            }
            S = Math.Pow(S / (masData.Length - 1), 1.0 / 2);
            scaleNumber = (C * N) / (S * Math.Pow(2 * Math.PI, 1.0 / 2));

        }
    }

    class TableColculations : RequiredValues, ICalculations
    {
        public double[,] masTable;
        public int n;
        public int m;
        double number;
        int numberStart;
        double count;
        public int v;
        public TableColculations(double[] masData)
        {
            this.masData = masData;
            base.Colculations();
            v = (int)k;
            n = (int)k + 2;
            m = 8;
            masTable = new double[n, m];
            number = C;
            numberStart = 0;
            count = 0;
            Colculations();

        }
        public TableColculations() { }

        public new void Colculations()
        {
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < m - 1; j++)
                {
                    if (j == 0)
                        masTable[i, j] = i;

                    if (j == 1)
                        masTable[i, j] = number;

                    if (j == 2 && i != 0)
                    {
                        number += C;
                        masTable[i, j] = number;
                    }
                    else if (j == 2 && i == 0)
                    {
                        masTable[i, j] = masData[0];
                        number = masData[0];
                    }

                    if (j == 3)
                        masTable[i, j] = (masTable[i, j - 1] + masTable[i, j - 2]) / 2;

                    if (j == 4)
                        masTable[i, j] = Math.Round(((masTable[i, j - 1] - averageX) / S), 4);

                    if (j == 5 && (i == 0 || i == n - 1))
                        masTable[i, j] = 0;
                    else if (j == 5)
                    {
                        for (int L = numberStart; L < masData.Length; L++)
                        {
                            if (masData[L] <= number)
                                count++;
                            else
                            {
                                numberStart = L;
                                break;
                            }
                        }
                        masTable[i, j] = count;
                        count = 0;
                    }

                    if (j == 6)
                        masTable[i, j] = Math.Round((scaleNumber * Math.Exp(-(1.0 / 2) * Math.Pow(masTable[i, j - 2], 2))), 4);

                }

            }
            for (int i = 0; i < n - 1; i++)
            {
                if (masTable[i, m - 3] < 5 || masTable[i + 1, m - 3] < 5)
                {
                    masTable[i, m - 1] = Math.Round((Math.Pow((masTable[i, m - 3] - masTable[i, m - 2] + masTable[i + 1, m - 3] - masTable[i + 1, m - 2]), 2) / (masTable[i, m - 2] + masTable[i + 1, m - 2])), 4);
                    i++;
                }
                else
                    masTable[i, m - 1] = Math.Round((Math.Pow(masTable[i, m - 3] - masTable[i, m - 2], 2) / masTable[i, m - 2]), 4);
            }

        }

        public double[,] GetMasTable()
        {
            return masTable;
        }
    }

    class CreateTable
    {
        int n;
        int m;
        public double sum;
        double[,] masTable;

        public CreateTable(double[,] masTable)
        {
            this.masTable = masTable;
            n = masTable.GetLength(0);
            m = masTable.GetLength(1);
        }

        public CreateTable() { }

        public void FillingInTheTable(DataGridView dataGridView, TextBox textBox)
        {
            dataGridView.RowCount = n;
            dataGridView.ColumnCount = m;
            int i, j;
            for (i = 0; i < n; i++)
            {
                for (j = 0; j < m; j++)
                {
                    dataGridView.Rows[i].Cells[j].Value = masTable[i, j];
                }
                sum += masTable[i, m - 1];
            }
            textBox.Text = Convert.ToString(sum);
        }
    }

    public delegate void Result(Label label);

    class OutPutResult
    {
        
        public event Result result = null;

        public void InvokeEvent(Label label)
        {
            result.Invoke(label);
        }
    }

    class CritRomanovsky
    {
        double critRomanovsky;
        
        public CritRomanovsky() { }

        public CritRomanovsky(TableColculations table, CreateTable createTable)
        {
            critRomanovsky = Math.Abs(createTable.sum - table.v) / Math.Pow(2 * table.v, 1.0 / 2);
        }

        public void ResultMetod(Label label)
        {
            if (critRomanovsky < 3)
                label.Text = "Гипотеза о нормальности выборочного ряда распределения не противоречит данным опыта.\n"+
                    "Таким образом, расхождение между выравнивающими и выборочными частотами можно считать не существенным\n"+
                    " и признать, что выборочное распределение подчиняется теоретическому закону, с которым его сравнивали.";
            else
                label.Text = "Расхождение между выравнивающими и выборочными частотами являются существенным можно\n" +
                    " признать, что выборочное распределение не подчиняется теоретическому закону, с которым его сравнивали.";
        }


    }
    class SaveTable
    {
        public void Save(DataGridView dataGritView, SaveFileDialog saveFileDialog)
        {

            saveFileDialog.Title = "Сохранить таблицу как ...";
            saveFileDialog.Filter = "Книга Excel|*.xlsx";
            saveFileDialog.AddExtension = true;
            saveFileDialog.FileName = "TableExcel_1";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet workSheet = workbook.ActiveSheet;

                    excelApp.Columns.ColumnWidth = 20;

                    for (int i = 1; i < dataGritView.Columns.Count + 1; i++)
                    {
                        excelApp.Cells[1, i] = dataGritView.Columns[i - 1].HeaderText;
                    }
                    for (int i = 0; i < dataGritView.Rows.Count; i++)
                    {
                        for (int j = 0; j < dataGritView.Columns.Count; j++)
                        {
                            excelApp.Cells[i + 2, j + 1] = dataGritView.Rows[i].Cells[j].Value;
                        }
                    }
                    excelApp.ActiveWorkbook.SaveCopyAs(saveFileDialog.FileName);
                    excelApp.ActiveWorkbook.Saved = true;
                    excelApp.Quit();
                MessageBox.Show("Таблица успешно сохранена!");
            }
                else
                MessageBox.Show("Таблица не была сохранена!");
            
        }
    }
    
   

    interface ICalculations
    {
        void Colculations();
    }

}