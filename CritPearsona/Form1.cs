using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;


namespace CritPearsona
{
    public partial class Form1 : Form
    {
        
        Data data = new WorkingWithData();
 
        public Form1()
        {
            InitializeComponent();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void label3_MouseEnter(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            pictureBox1.Visible = true;
        }

        private void label3_MouseLeave(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
               
            try
            {
                WorkingWithData data1 = new WorkingWithData(data.GetData(excelTable));
                data1.SortingData();
                if (data1.masData.Length != 0)
                    try
            {

                    TableColculations table = new TableColculations(data1.masData);
                    CreateTable createTable = new CreateTable(table.GetMasTable());

                    createTable.FillingInTheTable(dataGridView1, textBox1);

                    CritRomanovsky critRomanovsky = new CritRomanovsky(table, createTable);
                    OutPutResult outPutResult = new OutPutResult();
                    outPutResult.result += new Result(critRomanovsky.ResultMetod);
                    outPutResult.InvokeEvent(Result);
                }
                catch
                {
                    MessageBox.Show("Данные должны соответствовать примеру", "Неверный тип данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch
            {
                MessageBox.Show("Перед запуском откройте файл с данными", "Данных нет :(", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void button4_Click(object sender, EventArgs e)
        {           
            Close();
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 1)
                MessageBox.Show("Сначала откройте и заполните \nтаблицу нажав кнопку \"Запуск\"", "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                try
                {
                    SaveTable saveTable = new SaveTable();
                    saveTable.Save(dataGridView1, saveFileDialog1);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка сохранения таблицы!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
         
        }

        private double[,] excelTable;

        private int totalRows = 0;

        private int totalColumns = 0;

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
            try
            {
                DialogResult result = openFileDialog1.ShowDialog();
                if (result==DialogResult.OK)
                {
                   
                    ExcelPackage excelFile = new ExcelPackage(new FileInfo(openFileDialog1.FileName));

                    ExcelWorksheet workSheet = excelFile.Workbook.Worksheets[0];

                    totalColumns = workSheet.Dimension.End.Column;
                    totalRows = workSheet.Dimension.End.Row;

                    excelTable = new double[totalRows, totalColumns];

                    for(int rowIndex = 1; rowIndex<=totalRows;rowIndex++)
                    {
                        for (int columnsIndex = 1; columnsIndex <= totalColumns; columnsIndex++)
                        {
                            IEnumerable<string> row = workSheet.Cells[rowIndex, 1, rowIndex, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                            List<string> list = row.ToList<string>();

                            for (int i = 0; i < list.Count; i++)
                            {
                                excelTable[rowIndex - 1, i] = Convert.ToDouble(list[i].Replace('.', ','));
                            }
                        }
                    }
                   


                }
                else
                {
                    throw new Exception("Файл не выбран!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
    }
}
