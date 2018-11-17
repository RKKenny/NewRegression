using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SQLite;
using System.Globalization;
using System.Diagnostics;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        string filename = "";
        int n;
        int CompressionRateOverview;
        string ext;
        List<double> x = new List<double>();
        List<double> y = new List<double>();
        List<double> y2 = new List<double>();
        List<double> ye = new List<double>();
        List<double> ye2 = new List<double>();
        List<double> y3 = new List<double>();
        List<double> y4 = new List<double>();
        List<double> k1 = new List<double>();
        List<double> k2 = new List<double>();
        List<double> Ysh = new List<double>(); //производная 1
        List<double> Ysh2 = new List<double>(); //производная 2
        List<string> ss = new List<string>();
        //List<double> dInitialApp = new List<double>();
        List<double> kv = new List<double>(); //массив квадратичных отклонений
        bool Anchorr;
        bool Anchorr1; //чтобы не задействовать кнопки на графике


        List<double> p = new List<double>();     //массив параметров

        List<int> extr = new List<int>();  //массив точек, подозрительных на экстремум функции
        double eps, eps1, h, h1, d, IPS, T3, IPR, T5, mafx, FX30, FY30, mafy;     // точность, шаг, абсолютное значение шага
        int k;                         //кол-во итераций

        Excel.Application xlApp = new Excel.Application();
        List<double> Col1 = new List<double>();
        List<double> Col2 = new List<double>();

        List<double> Sel1 = new List<double>();
        List<double> Sel2 = new List<double>();

        double SelectionBorder1 = 0;
        double SelectionBorder2 = 0;
        int SelectionClicks = 1;

        double sumMean1;
        double sumMean2;

        double sumSKVO1;
        double sumSKVO2;

        double SKO;
        double MSE1;
        double MSE2;
        double disp;
        double VARI;
        int tb1c;
        int tb2c;
        double umnozh = 0;
        double umnozh2;
        double Student;
        double Pirson;

        List<string> RawDataFull = new List<string>();
        int CompressionRateSelection = 1;

        SQLiteConnection m_dbConnection;

        List<double> ssX = new List<double>();
        List<double> ssY = new List<double>();

        string ConnectionString = "Data Source = MyDatabase.sqlite; Version=3;";
        List<string> RawDataSelection = new List<string>();


        public Form1()
        {
            InitializeComponent();
            chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            chart2.ChartAreas[0].AxisX.ScaleView.Zoomable = true;

        }
        private void BtnLoadfromFile_Click(object sender, EventArgs e)
        {
            ext = "";            
                        
            openFileDialog1.Filter = "Поддерживаемые файлы(*.dat;*.txt)|*.dat;*.txt|Файлы DAT|*.dat|Текстовые файлы (*.txt)|*.txt";
            openFileDialog1.Title = "Выберите файл";

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filename = openFileDialog1.FileName;
            }
            else
            {
                return;
            }

            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            chart1.Series[2].Points.Clear();
            chart1.Series[3].Points.Clear();
            chart2.Series[0].Points.Clear();
            chart2.Series[1].Points.Clear();

            ext = Path.GetExtension(filename);

            if (ext == ".txt")
            {
                Anchorr1 = true;
                Anchorr = false;
                //ss.Clear();
                ss = new List<string>();
                ss = File.ReadAllLines(filename).ToList(); //построчно считываем данные из текстового файла в переменную  - массив строк
                n = ss.Count();

                dataGridView1.RowCount = 1;
                dataGridView1.ColumnCount = 7;
                dataGridView1[0, 0].Value = "№";
                dataGridView1[1, 0].Value = "x";
                dataGridView1[2, 0].Value = "y";
                dataGridView1[3, 0].Value = "F(x)";
                dataGridView1[4, 0].Value = "(Y - F(x)) ^ 2";
                dataGridView1[5, 0].Value = "F'";
                dataGridView1[6, 0].Value = "F''";

                ss.ForEach(delegate (String name)
                {
                    dataGridView1.Rows.Add(name.Split());
                });
                RawDataFull = ss;

                x = new List<double>();
                y = new List<double>();
                for (int i = 0; i < n; i++)
                {
                    x.Add(Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[1].Value.ToString()));
                    y.Add(Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[2].Value.ToString()));
                    chart1.Series[0].Points.AddXY(x[i], y[i]); //вывод графика по экспериментальным данным
                }
                Col2.Clear();
                Col2 = y;
                Col1.Clear();
                Col1 = x;

                //CompressionRateOverview = 1;
                //Col1 = CutArrayFromRows(RawDataFull, 0);
                //Col2 = CutArrayFromRows(RawDataFull, 1);

                MessageBox.Show("Выбран текстовый файл");
                BtnCalc.Enabled = true;
                
            }
            else
            {
                Anchorr1 = true;
                Anchorr = true;
                CompressionRateOverview = Convert.ToInt32(NmCompRate.Text);
                dataGridView1.Rows.Clear();
                x.Clear();
                y.Clear();

                if (!string.IsNullOrEmpty(filename) && File.Exists(filename))
                {
                    //извлечение ВСЕХ данных
                    RawDataFull = ExtractFromFile(filename, CompressionRateOverview);

                    Col1.Clear();
                    Col2.Clear();
                    //GLOBALS SHOULD BE ASSIGNED HERE
                    Col1 = CutArrayFromRows(RawDataFull, 0);
                    Col2 = CutArrayFromRows(RawDataFull, 1);

                    y = Col2;
                    x = Col1;
                    n = y.Count;

                    dataGridView1.RowCount = n + 1;
                    dataGridView1.ColumnCount = 7;
                    dataGridView1[0, 0].Value = "№";
                    dataGridView1[1, 0].Value = "x";
                    dataGridView1[2, 0].Value = "y";
                    dataGridView1[3, 0].Value = "F(x)";
                    dataGridView1[4, 0].Value = "(Y - F(x)) ^ 2";
                    dataGridView1[5, 0].Value = "F'";
                    dataGridView1[6, 0].Value = "F''";

                    for (int i = 0; i < n; i++)
                    {
                        dataGridView1[0, i + 1].Value = i + 1;
                        dataGridView1[1, i + 1].Value = Math.Round(x[i], 4);
                        dataGridView1[2, i + 1].Value = Math.Round(y[i], 4);

                    }
                    foreach (DataGridViewColumn column in dataGridView1.Columns)
                    {
                        column.SortMode = DataGridViewColumnSortMode.NotSortable;
                    }
                    chart1.Series[0].Points.Clear();
                    chart1.Series[1].Points.Clear();
                    chart1.Series[2].Points.Clear();
                    chart1.Series[3].Points.Clear();
                    chart2.Series[0].Points.Clear();
                    chart2.Series[1].Points.Clear();

                    BuildGraph2(false, true);
                    BtnCalc.Enabled = true;
                }                                
            }
            txtfortest.Text = n.ToString();
        }

        private void BuildGraph2(bool Selection, bool c2)
        {
            if (Selection)
            {
                if (c2) chart1.Series[0].Points.DataBindXY(Sel1, Sel2);
            }
            else
            {
                if (c2) chart1.Series[0].Points.DataBindXY(Col1, Col2);
            }
        }

        public List<string> ExtractFromFile(string filename, double compressionrate)
        {
            List<string> Extraction = new List<string>();
            Extraction.Clear();
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            try
            {
                //поиск первой пустой строки
                int rightindex = 1;
                Excel.Range RightCell;
                RightCell = (Excel.Range)xlWorkSheet.Cells[1, 1];

                //найти границу пустоты с точностью до 5000
                int SearchStep = 5000;
                while (RightCell.Value2 != null)
                {
                    rightindex += SearchStep;
                    RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
                }

                //найти точную границу методом половинного деления
                int leftindex = rightindex - SearchStep;
                Excel.Range LeftCell;
                LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
                while (SearchStep > 0)
                {
                    LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
                    RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
                    if ((LeftCell.Value2 != null) && (RightCell.Value2 == null))
                    {
                        SearchStep = SearchStep / 2;
                        leftindex += SearchStep;
                    }
                    if ((LeftCell.Value2 == null) && (RightCell.Value2 == null))
                        leftindex -= SearchStep;
                }

                Excel.Range xlRange = xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[1, 1], (Excel.Range)xlWorkSheet.Cells[leftindex + 1, 1]);

                int j = 0;

                //извлечение элементов из рейнджа в массив
                //Extraction = xlRange.Count / Convert.ToInt32(compressionrate) + 1;
                for (int i = 0; i < xlRange.Count; i += Convert.ToInt32(compressionrate))
                {
                    Extraction.Insert(j, (xlRange[i + 1, 1].Value2));
                    j++;
                }
                
                xlWorkBook.Close(0);
                xlApp.Quit();

                return Extraction;
            }
            catch (Exception e)
            {
                xlWorkBook.Close(0);
                xlApp.Quit();
                MessageBox.Show("Не удалось загрузить значения" + Environment.NewLine + e);
                throw;
            }
        }

        public List<double> CutArrayFromRows(List<string> data, int index)
        {
            var Col = new List<double>();
            Col.Clear();

            for (int i = 0; i < data.Count; i++)
                Col.Add((Convert.ToDouble(CutValuesFromStringDat(data[i])[index])));

            return Col;
        }

        private List<double> CutValuesFromStringDat(string s)
        {            
            //try
            //{
                List<double> Output = new List<double>();
                Output.Clear();
                int CutCount = 0;
                int PrevComma = -1;
                int NextComma = -1;
                string tempStr = "";

                while (CutCount < 11)
                {
                    NextComma = s.IndexOf(',', PrevComma + 1);

                    if ((CutCount == 0) || (CutCount == 1) || (CutCount == 4) || (CutCount == 7) || (CutCount == 10))
                    {
                        tempStr = s.Substring(PrevComma + 1, NextComma - PrevComma - 1).Replace('.', ',');

                        switch (CutCount)
                        {
                            case 0: Output.Insert(0, (Convert.ToDouble(tempStr))); break;
                            case 1: Output.Insert(1, (Convert.ToDouble(tempStr))); break;
                            case 4: Output.Insert(2, (Convert.ToDouble(tempStr))); break;
                            case 7: Output.Insert(3, (Convert.ToDouble(tempStr))); break;
                            case 10: Output.Insert(4, (Convert.ToDouble(tempStr))); break;
                        }

                        PrevComma = NextComma;
                        CutCount++;
                    }
                    else
                    {
                        PrevComma = NextComma;
                        CutCount++;
                    }
                }
                return Output;
            //}
            //catch (Exception e)
            //{
            //    xlApp.Quit();
            //    MessageBox.Show("Проверьте правильность границ" + Environment.NewLine + e);
            //    throw e;
            //}
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DBCreate();
            DBLoad();
        }

        private void DBCreate()
        {
            if (!File.Exists(GetApplicationExecutableDirectoryName() + ConstBackSlash + "MyDatabase.sqlite"))
            {
                MessageBox.Show("Первый запуск программы. Создаю базу данных" +
                    Environment.NewLine + "Если база данных уже существует и Вы видите это сообщение, то" +
                    Environment.NewLine + "немедленно сделайте бекап! (Папка с программой\\MyDatabase.sqlite)");
                //создание новой базы
                SQLiteConnection.CreateFile("MyDatabase.sqlite");

                //создание коннекшна
                m_dbConnection = new SQLiteConnection("Data Source=MyDatabase.sqlite;Version=3;");

                //открытие коннекшна
                m_dbConnection.Open();

                DBQueryExecute(@"
                CREATE TABLE Entry(
                id_entry integer PRIMARY KEY, 
                datemark varchar(20), 
                name varchar(60)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE Source(
                id_source integer PRIMARY KEY,
                id_entry integer, 
                x real, 
                y1 real,
                ye real, 
                k1 real, 
                k2 real,
                kv real,
                Ysh real, 
                Ysh2 real,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE InitApp(
                id_IA integer PRIMARY KEY,
                id_entry integer,
                a0 text,
                b0 text,
                c0 text,
                a1 text,
                b1 text,
                c1 text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE Koef(
                id_k integer PRIMARY KEY,
                id_entry integer,
                a0 text,
                b0 text,
                c0 text,
                a1 text,
                b1 text,
                c1 text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE AppEx(
                id_e integer PRIMARY KEY,
                id_entry integer,
                x1 text,
                y1 text,
                x2 text,
                y2 text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE YshEx(
                id_y1 integer PRIMARY KEY,
                id_entry integer,
                x1 text,
                y1 text,
                x2 text,
                y2 text,
                xm text,
                ym text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE Ysh2Ex(
                id_y2 integer PRIMARY KEY,
                id_entry integer,
                x1 text,
                y1 text,
                x2 text,
                y2 text,
                xm text,
                ym text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE Indexes(
                id_ind integer PRIMARY KEY,
                id_entry integer,
                IPS text,
                IPR text,
                YshD text,
                T3 text,
                A1 text,
                A2 text,
                A12 text,
                A21 text,
                Y30 text,
                VDM text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE Stats(
                id_s integer PRIMARY KEY,
                id_entry integer,
                Mean text,
                SKVO text,
                Disp text,
                SKO text,
                MSE text,
                Vari text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");
                DBQueryExecute(@"
                CREATE TABLE Addd(
                id_a integer PRIMARY KEY,
                id_entry integer,
                MeanErr text,
                SumKVO text,
                FOREIGN KEY(id_entry) REFERENCES Entry(id_entry)
                )
                ");

                //закрытие коннекшна
                m_dbConnection.Close();
                return;
            }            
            else
            {
                MessageBox.Show("База данных загружена");
            }

        }

        public string GetApplicationExecutableDirectoryName()
        {
            return Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
        }

        public const string ConstBackSlash = "\\";

        private void DBQueryExecute(string str)
        {
            SQLiteCommand command = new SQLiteCommand(str, m_dbConnection);
            command.ExecuteNonQuery();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            SaveToDB();
            DBLoad();
        }

        public void SaveToDB()
        {
            if (string.IsNullOrWhiteSpace(BDField.Text) == false)
            {
                if (dataGridView1.Rows.Count != 0 && dataGridView1.Rows != null)
                {
                    if (txtA0start.Text != "" && txtB0start.Text != "" && txtC0start.Text != ""
                    && txtA1start.Text != "" && txtB1start.Text != "" && txtC1start.Text != "")
                    {

                        if (txtA0kof.Text != "" && txtB0kof.Text != "" && txtC0kof.Text != ""
                                && txtA1kof.Text != "" && txtB1kof.Text != "" && txtC1kof.Text != "")
                        {

                            if (txty1X1m.Text != "" && txty1X2m.Text != "" && txty1Y1m.Text != ""
                                    && txty1Y2m.Text != "")
                            {
                                if (txtYshX1m.Text != "" && txtYshX2m.Text != "" && txtYshY1m.Text != ""
                                    && txtYshY2m.Text != "" && txtYshX1mi.Text != "" && txtYshY1mi.Text != "")
                                {
                                    if (txtYsh2X1m.Text != "" && txtYsh2X2m.Text != "" && txtYsh2Y1m.Text != ""
                                        && txtYsh2Y2m.Text != "" && txtYsh2X1mi.Text != "" && txtYsh2Y1mi.Text != "")
                                    {
                                        if (txtIndexA1.Text != "" && txtIndexA12Div.Text != "" && txtIndexA2.Text != ""
                                            && txtIndexA21Div.Text != "" && txtIndexIPR.Text != ""
                                            && txtIndexIPS.Text != "" && txtIndexT3.Text != ""
                                            && txtIndexVDM.Text != "" && txtIndexY30.Text != "" && txtIndexYshDiv.Text != "")
                                        {
                                            if (txtMean1.Text != "" && txtSKVO1.Text != "" && txtDisp.Text != ""
                                                && textSKO.Text != "" && txtMSE1.Text != "" && txtVar.Text != "")
                                            {
                                                if (TxtMeanErr.Text != "" && TxtSumSkvo.Text != "")
                                                {
                                                    //найти макс id в базе
                                                    int MaxID = 0;

                                                    Boolean EntriesExist = false;

                                                    string sql = "select count(id_entry) as Rows FROM Entry";
                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            using (SQLiteDataReader rdr = cmd.ExecuteReader())
                                                            {
                                                                if (rdr.Read())
                                                                    if (Convert.ToInt32(rdr["Rows"]) > 0)
                                                                        EntriesExist = true;
                                                                c.Close();
                                                            }
                                                        }
                                                    }

                                                    if (EntriesExist)
                                                    {
                                                        sql = "select max(id_entry) AS max FROM Entry";
                                                        using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                        {
                                                            c.Open();
                                                            using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                            {
                                                                using (SQLiteDataReader rdr = cmd.ExecuteReader())
                                                                {
                                                                    if (rdr.Read())
                                                                        MaxID = Convert.ToInt32(rdr["max"]);
                                                                    c.Close();
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else MaxID = 0;

                                                    MaxID++;
                                                    string CurrentDate = DateTime.Now.ToString();
                                                    string Name = BDField.Text;

                                                    sql = @"insert into Entry
                                            Values(NULL, @date, @name)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@date", CurrentDate);
                                                            cmd.Parameters.AddWithValue("@name", Name);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }


                                                    //ЗАПИСЬ СОУСА В БАЗУ 
                                                    for (int i = 0; i < x.Count; i++)
                                                    {
                                                        sql = @"INSERT INTO Source
                                                Values (NULL, @id2, @x, @y1, @ye, @k1, @k2, @kv, @Ysh, @Ysh2)";

                                                        using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                        {
                                                            c.Open();
                                                            using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                            {
                                                                cmd.Parameters.AddWithValue("@id2", MaxID);
                                                                cmd.Parameters.AddWithValue("@x", x[i]);
                                                                cmd.Parameters.AddWithValue("@y1", y[i]);
                                                                cmd.Parameters.AddWithValue("@ye", ye[i]);
                                                                cmd.Parameters.AddWithValue("@k1", k1[i]);
                                                                cmd.Parameters.AddWithValue("@k2", k2[i]);
                                                                cmd.Parameters.AddWithValue("@kv", kv[i]);
                                                                cmd.Parameters.AddWithValue("@Ysh", Ysh[i]);
                                                                cmd.Parameters.AddWithValue("@Ysh2", Ysh2[i]);
                                                                cmd.ExecuteNonQuery();
                                                            }
                                                            c.Close();
                                                        }
                                                    }

                                                    sql = @"insert into InitApp
                                            Values(NULL, @id2, @a0, @b0, @c0, @a1, @b1, @c1)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@a0", txtA0start.Text);
                                                            cmd.Parameters.AddWithValue("@b0", txtB0start.Text);
                                                            cmd.Parameters.AddWithValue("@c0", txtC0start.Text);
                                                            cmd.Parameters.AddWithValue("@a1", txtA1start.Text);
                                                            cmd.Parameters.AddWithValue("@b1", txtB1start.Text);
                                                            cmd.Parameters.AddWithValue("@c1", txtC1start.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into Koef
                                            Values(NULL, @id2, @a0, @b0, @c0, @a1, @b1, @c1)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@a0", txtA0kof.Text);
                                                            cmd.Parameters.AddWithValue("@b0", txtB0kof.Text);
                                                            cmd.Parameters.AddWithValue("@c0", txtC0kof.Text);
                                                            cmd.Parameters.AddWithValue("@a1", txtA1kof.Text);
                                                            cmd.Parameters.AddWithValue("@b1", txtB1kof.Text);
                                                            cmd.Parameters.AddWithValue("@c1", txtC1kof.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into AppEx
                                            Values(NULL, @id2, @x1, @y1, @x2, @y2)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@x1", txty1X1m.Text);
                                                            cmd.Parameters.AddWithValue("@y1", txty1Y1m.Text);
                                                            cmd.Parameters.AddWithValue("@x2", txty1X2m.Text);
                                                            cmd.Parameters.AddWithValue("@y2", txty1Y2m.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into YshEx
                                            Values(NULL, @id2, @x1, @y1, @x2, @y2, @xm, @ym)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@x1", txtYshX1m.Text);
                                                            cmd.Parameters.AddWithValue("@y1", txtYshY1m.Text);
                                                            cmd.Parameters.AddWithValue("@x2", txtYshX2m.Text);
                                                            cmd.Parameters.AddWithValue("@y2", txtYshY2m.Text);
                                                            cmd.Parameters.AddWithValue("@xm", txtYshX1mi.Text);
                                                            cmd.Parameters.AddWithValue("@ym", txtYshY1mi.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into Ysh2Ex
                                            Values(NULL, @id2, @x1, @y1, @x2, @y2, @xm, @ym)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@x1", txtYsh2X1m.Text);
                                                            cmd.Parameters.AddWithValue("@y1", txtYsh2Y1m.Text);
                                                            cmd.Parameters.AddWithValue("@x2", txtYsh2X2m.Text);
                                                            cmd.Parameters.AddWithValue("@y2", txtYsh2Y2m.Text);
                                                            cmd.Parameters.AddWithValue("@xm", txtYsh2X1mi.Text);
                                                            cmd.Parameters.AddWithValue("@ym", txtYsh2Y1mi.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into Indexes
                                            Values(NULL, @id2, @IPS, @IPR, @YshD, @T3, @A1, @A2, @A12, @A21, @Y30, @VDM)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@IPS", txtIndexIPS.Text);
                                                            cmd.Parameters.AddWithValue("@IPR", txtIndexIPR.Text);
                                                            cmd.Parameters.AddWithValue("@YshD", txtIndexYshDiv.Text);
                                                            cmd.Parameters.AddWithValue("@T3", txtIndexT3.Text);
                                                            cmd.Parameters.AddWithValue("@A1", txtIndexA1.Text);
                                                            cmd.Parameters.AddWithValue("@A2", txtIndexA2.Text);
                                                            cmd.Parameters.AddWithValue("@A12", txtIndexA12Div.Text);
                                                            cmd.Parameters.AddWithValue("@A21", txtIndexA21Div.Text);
                                                            cmd.Parameters.AddWithValue("@Y30", txtIndexY30.Text);
                                                            cmd.Parameters.AddWithValue("@VDM", txtIndexVDM.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into Stats
                                            Values(NULL, @id2, @Mean, @SKVO, @Disp, @SKO, @MSE, @Vari)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@Mean", txtMean1.Text);
                                                            cmd.Parameters.AddWithValue("@SKVO", txtSKVO1.Text);
                                                            cmd.Parameters.AddWithValue("@Disp", txtDisp.Text);
                                                            cmd.Parameters.AddWithValue("@SKO", textSKO.Text);
                                                            cmd.Parameters.AddWithValue("@MSE", txtMSE1.Text);
                                                            cmd.Parameters.AddWithValue("@Vari", txtVar.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    sql = @"insert into Addd
                                            Values(NULL, @id2, @MeanErr, @SumKVO)";

                                                    using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                                                    {
                                                        c.Open();
                                                        using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                                                        {
                                                            cmd.Parameters.AddWithValue("@id2", MaxID);
                                                            cmd.Parameters.AddWithValue("@MeanErr", TxtMeanErr.Text);
                                                            cmd.Parameters.AddWithValue("@SumKVO", TxtSumSkvo.Text);
                                                            cmd.ExecuteNonQuery();
                                                        }
                                                        c.Close();
                                                    }

                                                    MessageBox.Show("Измерение " + Name + " сохранено в базе данных!");

                                                    return;
                                                }
                                                else
                                                {
                                                    tabControl1.SelectTab(TbP1);
                                                    TxtMeanErr.BackColor = Color.Yellow;
                                                    TxtSumSkvo.BackColor = Color.Yellow;
                                                    MessageBox.Show("Отсутствуют дополнительные данные");
                                                    TxtMeanErr.BackColor = Color.Transparent;
                                                    TxtSumSkvo.BackColor = Color.Transparent;
                                                    return;
                                                }
                                            }
                                            else
                                            {
                                                tabControl1.SelectTab(TbP5);
                                                txtMean1.BackColor = Color.Yellow;
                                                txtSKVO1.BackColor = Color.Yellow;
                                                txtDisp.BackColor = Color.Yellow;
                                                textSKO.BackColor = Color.Yellow;
                                                txtMSE1.BackColor = Color.Yellow;
                                                txtVar.BackColor = Color.Yellow;
                                                MessageBox.Show("Отсутствуют данные о статистике");
                                                txtMean1.BackColor = SystemColors.Window;
                                                txtSKVO1.BackColor = SystemColors.Window;
                                                txtDisp.BackColor = SystemColors.Window;
                                                textSKO.BackColor = SystemColors.Window;
                                                txtMSE1.BackColor = SystemColors.Window;
                                                txtVar.BackColor = SystemColors.Window;
                                                return;
                                            }
                                        }
                                        else
                                        {
                                            tabControl1.SelectTab(TbP2);
                                            groupBox8.BackColor = Color.Yellow;
                                            MessageBox.Show("Отсутствуют данные об индексах");
                                            groupBox8.BackColor = Color.Transparent;
                                            return;
                                        }

                                    }
                                    else
                                    {
                                        tabControl1.SelectTab(TbP2);
                                        groupBox4.BackColor = Color.Yellow;
                                        MessageBox.Show("Отсутствуют данные об экстремумах второй производной");
                                        groupBox4.BackColor = Color.Transparent;
                                        return;
                                    }
                                }
                                else
                                {
                                    tabControl1.SelectTab(TbP2);
                                    groupBox5.BackColor = Color.Yellow;
                                    MessageBox.Show("Отсутствуют данные об экстремумах первой производной");
                                    groupBox5.BackColor = Color.Transparent;
                                    return;
                                }
                            }
                            else
                            {
                                tabControl1.SelectTab(TbP2);
                                groupBox3.BackColor = Color.Yellow;
                                MessageBox.Show("Отсутствуют данные об экстремумах аппроксимирующей функции");
                                groupBox3.BackColor = Color.Transparent;
                                return;
                            }
                        }
                        else
                        {
                            tabControl1.SelectTab(TbP1);
                            groupBox1.BackColor = Color.Yellow;
                            MessageBox.Show("Отсутствуют коэффициенты");
                            groupBox1.BackColor = Color.Transparent;
                            return;
                        }
                    }
                    else
                    {
                        tabControl1.SelectTab(TbP1);
                        groupBox2.BackColor = Color.Yellow;
                        MessageBox.Show("Отсутствует начальное приближение");
                        groupBox2.BackColor = Color.Transparent;
                        return;
                    }
                }
                else
                {
                    tabControl1.SelectTab(TbP1);
                    dataGridView1.BackgroundColor = Color.Yellow;
                    MessageBox.Show("Внимание! Таблица со значениями пуста");
                    dataGridView1.BackgroundColor = SystemColors.AppWorkspace;
                    return;
                }
            }
            else
            {
                tabControl1.SelectTab(TbP1);
                BDField.BackColor = Color.Yellow;
                MessageBox.Show("Введите имя для базы данных");
                BDField.BackColor = SystemColors.Window;
                return;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(TbP4);
            DBLoad();            
        }

        public void DBLoad()
        {
            //найти число строк в базе
            int MaxID = 0;
            string sql = "select count(id_entry) AS Ids FROM Entry";
            using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
            {
                c.Open();
                using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                {
                    using (SQLiteDataReader rdr = cmd.ExecuteReader())
                    {
                        if (rdr.Read())
                            MaxID = Convert.ToInt32(rdr["Ids"]);
                    }
                }
                c.Close();
            }

            if (MaxID != 0)
            {
                //вывести всё в ДГВ из Entry
                dataGridView3.RowCount = MaxID;
                dataGridView3.ColumnCount = 2;

                int i = 0;
                sql = "select * from Entry";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                dataGridView3[0, i].Value = Convert.ToString(rdr["id_entry"]);
                                dataGridView3[1, i].Value = Convert.ToString((rdr["datemark"]), CultureInfo.CreateSpecificCulture("ru-RU")) + " " + Convert.ToString(rdr["name"]);
                                i++;
                            }
                        }
                    }
                    c.Close();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                int SelectedId = Convert.ToInt32(dataGridView3[0, dataGridView3.CurrentCell.RowIndex].Value);

                x.Clear();
                y.Clear();
                ye.Clear();
                k1.Clear();
                k2.Clear();
                kv.Clear();
                Ysh.Clear();
                Ysh2.Clear();

                string sql = "select * from Entry where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                BDField.Text = (Convert.ToString(rdr["name"]));
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Source where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                x.Add(Convert.ToDouble(rdr["x"], CultureInfo.InvariantCulture));
                                y.Add(Convert.ToDouble(rdr["y1"], CultureInfo.InvariantCulture));
                                ye.Add(Convert.ToDouble(rdr["ye"], CultureInfo.InvariantCulture));
                                k1.Add(Convert.ToDouble(rdr["k1"], CultureInfo.InvariantCulture));
                                k2.Add(Convert.ToDouble(rdr["k2"], CultureInfo.InvariantCulture));
                                kv.Add(Convert.ToDouble(rdr["kv"], CultureInfo.InvariantCulture));
                                Ysh.Add(Convert.ToDouble(rdr["Ysh"], CultureInfo.InvariantCulture));
                                Ysh2.Add(Convert.ToDouble(rdr["Ysh2"], CultureInfo.InvariantCulture));
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from InitApp where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtA0start.Text = Convert.ToString(rdr["a0"]);
                                txtB0start.Text = Convert.ToString(rdr["b0"]);
                                txtC0start.Text = Convert.ToString(rdr["c0"]);
                                txtA1start.Text = Convert.ToString(rdr["a1"]);
                                txtB1start.Text = Convert.ToString(rdr["b1"]);
                                txtC1start.Text = Convert.ToString(rdr["c1"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Koef where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtA0kof.Text = Convert.ToString(rdr["a0"]);
                                txtB0kof.Text = Convert.ToString(rdr["b0"]);
                                txtC0kof.Text = Convert.ToString(rdr["c0"]);
                                txtA1kof.Text = Convert.ToString(rdr["a1"]);
                                txtB1kof.Text = Convert.ToString(rdr["b1"]);
                                txtC1kof.Text = Convert.ToString(rdr["c1"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from AppEx where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txty1X1m.Text = Convert.ToString(rdr["x1"]);
                                txty1Y1m.Text = Convert.ToString(rdr["y1"]);
                                txty1X2m.Text = Convert.ToString(rdr["x2"]);
                                txty1Y2m.Text = Convert.ToString(rdr["y2"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from YshEx where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtYshX1m.Text = Convert.ToString(rdr["x1"]);
                                txtYshY1m.Text = Convert.ToString(rdr["y1"]);
                                txtYshX2m.Text = Convert.ToString(rdr["x2"]);
                                txtYshY2m.Text = Convert.ToString(rdr["y2"]);
                                txtYshX1mi.Text = Convert.ToString(rdr["xm"]);
                                txtYshY1mi.Text = Convert.ToString(rdr["ym"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Ysh2Ex where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtYsh2X1m.Text = Convert.ToString(rdr["x1"]);
                                txtYsh2Y1m.Text = Convert.ToString(rdr["y1"]);
                                txtYsh2X2m.Text = Convert.ToString(rdr["x2"]);
                                txtYsh2Y2m.Text = Convert.ToString(rdr["y2"]);
                                txtYsh2X1mi.Text = Convert.ToString(rdr["xm"]);
                                txtYsh2Y1mi.Text = Convert.ToString(rdr["ym"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Indexes where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtIndexIPS.Text = Convert.ToString(rdr["IPS"]);
                                txtIndexIPR.Text = Convert.ToString(rdr["IPR"]);
                                txtIndexYshDiv.Text = Convert.ToString(rdr["YshD"]);
                                txtIndexT3.Text = Convert.ToString(rdr["T3"]);
                                txtIndexA1.Text = Convert.ToString(rdr["A1"]);
                                txtIndexA2.Text = Convert.ToString(rdr["A2"]);
                                txtIndexA12Div.Text = Convert.ToString(rdr["A12"]);
                                txtIndexA21Div.Text = Convert.ToString(rdr["A21"]);
                                txtIndexY30.Text = Convert.ToString(rdr["Y30"]);
                                txtIndexVDM.Text = Convert.ToString(rdr["VDM"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Stats where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtMean1.Text = Convert.ToString(rdr["Mean"]);
                                txtSKVO1.Text = Convert.ToString(rdr["SKVO"]);
                                txtDisp.Text = Convert.ToString(rdr["Disp"]);
                                textSKO.Text = Convert.ToString(rdr["SKO"]);
                                txtMSE1.Text = Convert.ToString(rdr["MSE"]);
                                txtVar.Text = Convert.ToString(rdr["Vari"]);
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Addd where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                TxtMeanErr.Text = Convert.ToString(rdr["MeanErr"]);
                                TxtSumSkvo.Text = Convert.ToString(rdr["SumKVO"]);
                            }
                        }
                    }
                    c.Close();
                }

                chart1.Series[0].Points.Clear();
                chart1.Series[1].Points.Clear();
                chart1.Series[2].Points.Clear();
                chart1.Series[3].Points.Clear();
                chart2.Series[0].Points.Clear();
                chart2.Series[1].Points.Clear();
                //chart1.Series[7].Points.Clear();
                chart2.Visible = true;
                for (int i = 0; i < ye.Count; i++)
                {
                    //вывод графиков экспериментальных и аппроксимирующих значений функции
                    chart1.Series[0].Points.AddXY(x[i], y[i]);
                    chart1.Series[1].Points.AddXY(x[i], ye[i]);
                    chart1.Series[2].Points.AddXY(x[i], k1[i]);
                    chart1.Series[3].Points.AddXY(x[i], k2[i]);
                    chart2.Series[0].Points.AddXY(x[i], Ysh[i]);
                    chart2.Series[1].Points.AddXY(x[i], Ysh2[i]);
                }

                //заполнение дата грид вью

                n = x.Count;

                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();

                dataGridView1.RowCount = n + 1;
                dataGridView1.ColumnCount = 7;
                dataGridView1[0, 0].Value = "№";
                dataGridView1[1, 0].Value = "x";
                dataGridView1[2, 0].Value = "y";
                dataGridView1[3, 0].Value = "F(x)";
                dataGridView1[4, 0].Value = "(Y - F(x)) ^ 2";
                dataGridView1[5, 0].Value = "F'";
                dataGridView1[6, 0].Value = "F''";

                //textBox28.Text = Convert.ToString( x.Length);
                for (int i = 0; i < x.Count; i++)
                {
                    //textBox2.Text += ssX + Environment.NewLine;
                    dataGridView1[0, i + 1].Value = i + 1;
                    dataGridView1[1, i + 1].Value = Math.Round(x[i], 4);
                    dataGridView1[2, i + 1].Value = Math.Round(y[i], 4);
                    dataGridView1[3, i + 1].Value = Math.Round(ye[i], 4);
                    dataGridView1[4, i + 1].Value = Math.Round(kv[i], 4);
                    dataGridView1[5, i + 1].Value = Math.Round(Ysh[i], 4);
                    dataGridView1[6, i + 1].Value = Math.Round(Ysh2[i], 4);

                }
                txtfortest.Text = n.ToString();

                dataGridView4.Rows.Clear();
                dataGridView4.Refresh();

                dataGridView4.RowCount = ye.Count + 1;
                dataGridView4.ColumnCount = 4;
                dataGridView4[0, 0].Value = "№";
                dataGridView4[1, 0].Value = "F(x)";
                dataGridView4[2, 0].Value = "M-F(x)";
                dataGridView4[3, 0].Value = "(M-F(x))^2";


                for (int i = 0; i < ye.Count; i++)
                {
                    dataGridView4.Rows[i + 1].Cells[0].Value = i + 1;
                    dataGridView4.Rows[i + 1].Cells[1].Value = Math.Round(ye[i], 4);
                    dataGridView4.Rows[i + 1].Cells[2].Value = Math.Round(Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[1].Value) - Convert.ToDouble(txtMean1.Text), 4);
                    dataGridView4.Rows[i + 1].Cells[3].Value = Math.Round(Math.Pow((Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[1].Value) - Convert.ToDouble(txtMean1.Text)), 2), 4);
                }
                Anchorr1 = false;
                Anchorr = true;
                Sel1 = x;
                Sel2 = y;

                //активировать главную вкладку
                BtnCalc.Enabled = true;
                tabControl1.SelectTab(TbP1);
                btnStat.Enabled = true;
                BtnStatCalc.Enabled = true;
                button5.Enabled = false;
                button6.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Невозможно загрузить запись из базы данных." + Environment.NewLine +
                    "Либо она не выделена в таблице, либо произошла иная ошибка" + Environment.NewLine +
                    ex.Message);
            }
        }

        public void chart1_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (Anchorr1 == true)
                {
                    if (Anchorr == true)
                    {
                        //double datapoint = Convert.ToDouble(chart1.ChartAreas[0].AxisX.GetPosition(50));
                        double ClickPos = Convert.ToDouble(chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.X));

                        if (ClickPos < Col1.Min()) ClickPos = Col1.Min();
                        else if (ClickPos > Col1.Max()) ClickPos = Col1.Max();

                        SelectionClicks++;
                        button5.Enabled = true;
                        if (SelectionClicks % 2 == 1) SelectionBorder1 = ClickPos;
                        else SelectionBorder2 = ClickPos;
                        //label3.Text = Convert.ToString(chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Y));

                        if (SelectionBorder1 > SelectionBorder2)
                        {
                            SelectionBorder1 = SelectionBorder1 - SelectionBorder2;
                            SelectionBorder2 += SelectionBorder1;
                            SelectionBorder1 = SelectionBorder2 - SelectionBorder1;
                        }

                        if (SelectionBorder1 * SelectionBorder2 == 0)
                            button5.Enabled = true;

                        //ShowSelectionBorder(SelectionBorder1, SelectionBorder2);

                        label1.Text = "ЛЕВАЯ ГРАНИЦА: " + Convert.ToString(Math.Round(SelectionBorder1, 3)) + ", ПРАВАЯ ГРАНИЦА: " + Convert.ToString(Math.Round(SelectionBorder2, 3));
                    }
                    else if (Anchorr == false)
                    {
                        //double datapoint = Convert.ToDouble(chart1.ChartAreas[0].AxisX.GetPosition(50));
                        double ClickPos = Convert.ToDouble(chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.X));

                        if (ClickPos < Col1.Min()) ClickPos = Col1.Min();
                        else if (ClickPos > Col1.Max()) ClickPos = Col1.Max();

                        SelectionClicks++;
                        //button5.Enabled = true;
                        if (SelectionClicks % 2 == 1) SelectionBorder1 = ClickPos;
                        else SelectionBorder2 = ClickPos;
                        //label3.Text = Convert.ToString(chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Y));

                        if (SelectionBorder1 > SelectionBorder2)
                        {
                            SelectionBorder1 = SelectionBorder1 - SelectionBorder2;
                            SelectionBorder2 += SelectionBorder1;
                            SelectionBorder1 = SelectionBorder2 - SelectionBorder1;
                        }

                        //if (SelectionBorder1 * SelectionBorder2 == 0)
                        //button5.Enabled = true;

                        //ShowSelectionBorder(SelectionBorder1, SelectionBorder2);

                        label1.Text = "ЛЕВАЯ ГРАНИЦА: " + Convert.ToString(Math.Round(SelectionBorder1, 3)) + ", ПРАВАЯ ГРАНИЦА: " + Convert.ToString(Math.Round(SelectionBorder2, 3));
                    }
                }
            }
            catch(Exception ex)
            {
                xlApp.Quit();
                MessageBox.Show("Убедитесь, что график отображается и границы заданы правильно." +
                    Environment.NewLine + "Невозможно увеличить или задать границы графика в случае, " +
                    "если данные загружены из базы данных." +Environment.NewLine + ex);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            button6.Enabled = true;
            List<int> Borders = new List<int>();
            Borders.Clear();
            Borders = FindSelectionBorders(filename, SelectionBorder1, SelectionBorder2);
            //MessageBox.Show(Convert.ToString(Borders[0]) + ", " + Convert.ToString(Borders[1]));
            SelectionBorder1 = Borders[0];
            SelectionBorder2 = Borders[1];
            RawDataSelection = ExtractFromFile(filename, CompressionRateSelection, SelectionBorder1, SelectionBorder2);

            //Sel1.Clear();
            Sel1 = CutArrayFromRows(RawDataSelection, 0);
            //Sel2.Clear();
            Sel2 = CutArrayFromRows(RawDataSelection, 1);

            //присовение глобальным массивам новых значений
            y = Sel2;
            x = Sel1;


            //сглаживание массива иксов
            List<double> x2 = new List<double>(x.Count);
            //x2.Clear();
            x2 = SmoothArray(x);
            //x.Clear();
            x = x2;
            n = x.Count;

            BuildGraph2(true, true);

            dataGridView1.Rows.Clear();
            dataGridView1.RowCount = x.Count + 1;
            dataGridView1.ColumnCount = 7;
            dataGridView1[0, 0].Value = "№";
            dataGridView1[1, 0].Value = "x";
            dataGridView1[2, 0].Value = "y";
            dataGridView1[3, 0].Value = "F(x)";
            dataGridView1[4, 0].Value = "(Y - F(x)) ^ 2";
            dataGridView1[5, 0].Value = "F'";
            dataGridView1[6, 0].Value = "F''";

            for (int i = 0; i < x.Count; i++)
            {
                dataGridView1.Rows[i + 1].Cells[0].Value = i + 1;
                dataGridView1[1, i + 1].Value = Math.Round(x[i], 4);
                dataGridView1[2, i + 1].Value = Math.Round(y[i], 4);
                //chart1.Series[0].Points.AddXY(x[i], y[i]); //вывод графика по экспериментальным данным
            }
            txtfortest.Text = n.ToString();
        }

        private List<double> SmoothArray(List<double> Arr)
        {
            List<double> Arr2 = new List<double>();
            Arr2.Clear();
            //Arr2 = Arr;

            int i = 0;
            double CurrentValue = 0;
            double NextValue = 0;
            int CurrentIndex = 0;
            int NextIndex = 0;

            while (CurrentIndex < Arr.Count - 1)
            {
                CurrentValue = Arr[CurrentIndex];
                while ((Arr[NextIndex] == CurrentValue) && (NextIndex < Arr.Count - 1))
                    NextIndex++;

                NextValue = Arr[NextIndex];

                //сгладить фрагмент массива
                for (i = CurrentIndex; i < NextIndex; i++)
                {
                    Arr2.Add(CurrentValue + (i - CurrentIndex) * (NextValue - CurrentValue) / (NextIndex - CurrentIndex));
                }

                CurrentIndex = NextIndex;
            }

            return Arr2;
        }

        public List<string> ExtractFromFile(string filename, double cRate, double left, double right)
        {
            List<string> Extraction = new List<string>();

            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range xlRange = (Excel.Range)xlWorkSheet.get_Range((Excel.Range)xlWorkSheet.Cells[left, 1], (Excel.Range)xlWorkSheet.Cells[right, 1]);
            //извлечение элементов из рейнджа в массив
            //Extraction = new string[xlRange.Count / Convert.ToInt32(cRate)];
            for (int i = 0; i < xlRange.Count; i += Convert.ToInt32(cRate))
            {
                //Extraction[i / Convert.ToInt32(cRate)] = xlRange[i, 1].Value2;
                Extraction.Insert((i / Convert.ToInt32(cRate)), xlRange[i, 1].Value2);
            }

            //GetExcelProcess(xlApp).Kill();
            xlWorkBook.Close(0);
            xlApp.Quit();
            return Extraction;
        }

        private List<int> FindSelectionBorders(string filename, double left, double right)
        {
            List<int> Borders1 = new List<int>();
            //Borders1.Clear();
            //Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;

            //xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(@filename, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            //найти LEFT и RIGHT по имеющимся таймингам
            int SearchStep = 5000;
            int leftindex = 1; int rightindex = leftindex;
            Excel.Range CurrentCell;
            CurrentCell = (Excel.Range)xlWorkSheet.Cells[1 + SearchStep, 1];
            Excel.Range PrevCell;
            PrevCell = (Excel.Range)xlWorkSheet.Cells[1, 1];
            double CurValue = CutValuesFromStringDat(CurrentCell.Value2)[0];
            double PrevValue = CutValuesFromStringDat(PrevCell.Value2)[0];

            //найти примерные границы левого края
            while ((((CurValue - left) * (PrevValue - left)) > 0) || !(string.IsNullOrEmpty(CurrentCell.Value2)))
            {
                leftindex += SearchStep;
                CurrentCell = (Excel.Range)xlWorkSheet.Cells[leftindex + SearchStep, 1];
                PrevCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];

                if (string.IsNullOrEmpty(CurrentCell.Value2))
                    break;
                else
                {
                    CurValue = CutValuesFromStringDat(CurrentCell.Value2)[0];
                    PrevValue = CutValuesFromStringDat(PrevCell.Value2)[0];
                }
            }

            rightindex = leftindex + SearchStep;

            //найти точную границу методом половинного деления
            Excel.Range LeftCell;
            Excel.Range RightCell;
            LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
            RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
            while (SearchStep > 0)
            {
                LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
                if ((LeftCell.Value2 != null) && (RightCell.Value2 == null))
                {
                    SearchStep = SearchStep / 2;
                    rightindex -= SearchStep;
                }
                if ((LeftCell.Value2 != null) && (RightCell.Value2 != null))
                    rightindex += SearchStep;
            }

            rightindex--;
            int RightNotEmpty = rightindex;

            //найденная выше граница является ГРАНИЦЕЙ ЗНАЧЕНИЙ, ГДЕ ВООБЩЕ В ПРИНЦИПЕ ЕСТЬ КАКИЕ-ТО ЗНАЧЕНИЯ
            //теперь необходимо найти левую границу ИСКОМОГО ДИАПАЗОНА
            //LEFT = 5000, RIGHT = 9570

            SearchStep = RightNotEmpty - leftindex;
            Excel.Range MidCell;
            MidCell = (Excel.Range)xlWorkSheet.Cells[leftindex];

            //поиск левой точной границы
            while (Math.Abs(CutValuesFromStringDat(LeftCell.Value2)[0] - left) > 0.1)
            {
                double cval = CutValuesFromStringDat(LeftCell.Value2)[0];
                SearchStep = Convert.ToInt32(80 * (left - cval));
                leftindex += SearchStep;
                LeftCell = (Excel.Range)xlWorkSheet.Cells[leftindex, 1];
            }

            RightCell = (Excel.Range)xlWorkSheet.Cells[RightNotEmpty, 1];
            //поиск правой точной границы
            while (Math.Abs(CutValuesFromStringDat(RightCell.Value2)[0] - right) > 0.1)
            {
                double cval = CutValuesFromStringDat(RightCell.Value2)[0];
                SearchStep = Convert.ToInt32(80 * (cval - right));
                rightindex -= SearchStep;
                RightCell = (Excel.Range)xlWorkSheet.Cells[rightindex, 1];
            }

            xlWorkBook.Close();
            xlApp.Quit();
            Borders1.Insert(0, leftindex);
            Borders1.Insert(1, rightindex);
            return Borders1;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            BuildGraph2(false, true);
            button5.Enabled = false;
            button6.Enabled = false;
        }

        private void BtnCalc_Click(object sender, EventArgs e)
        {
            //try
            //{

            chart1.Visible = true;
            chart2.Visible = false;

            textBox28.Text = "";
            textBox29.Text = "";

            InitialApproximation();
            
            List<double> kv1 = new List<double>();
            List<double> kv2 = new List<double>();
            kv.Clear();
            kv1.Clear();
            kv2.Clear();
            //начальное приближение
            p.Clear();
            p.Add(Convert.ToDouble(txtA0start.Text.Trim()));
            p.Add(Convert.ToDouble(txtB0start.Text.Trim()));
            p.Add(Convert.ToDouble(txtC0start.Text.Trim()));
            p.Add(Convert.ToDouble(txtA1start.Text.Trim()));
            p.Add(Convert.ToDouble(txtB1start.Text.Trim()));
            p.Add(Convert.ToDouble(txtC1start.Text.Trim()));
            //расчет методом покоординатного спуска
            eps = 0.001; //погрешность
            h = 0.005; //шаг поиска - начальное значение
            k = 50; //кол-во итераций
            eps1 = eps / k;
            do
            {
                d = Math.Abs(h);
                for (int i = 0; i < 6; i++)
                {
                    h1 = h;
                    scan(i);
                }
                h = h / k;
            }
            while (d > eps);

            ye = F(x, p);
            for (int i = 0; i < n; i++)
                kv.Insert(i, ((y[i] - y2[i]) * (y[i] - y2[i])));
            k1 = Fk1(x, p);
            for (int i = 0; i < n; i++)
                kv1.Insert(i, (y[i] - k1[i]) * (y[i] - k1[i]));
            k2 = Fk2(x, p);
            for (int i = 0; i < n; i++)
                kv2.Insert(i, ((y[i] - k2[i]) * (y[i] - k2[i])));

            //вывод расчетных коэф-тов функции
            txtA0kof.Text = Math.Round(p[0], 6).ToString();
            txtB0kof.Text = Math.Round(p[1], 6).ToString();
            txtC0kof.Text = Math.Round(p[2], 6).ToString();
            txtA1kof.Text = Math.Round(p[3], 6).ToString();
            txtB1kof.Text = Math.Round(p[4], 6).ToString();
            txtC1kof.Text = Math.Round(p[5], 6).ToString();

            Ysh = Fsh(x, p);

            Ysh2 = Fsh2(x, p);

            //вывод расчетных значений функции и квадр.отклонений
            chart1.Series[0].Points.Clear();
            chart1.Series[1].Points.Clear();
            chart1.Series[2].Points.Clear();
            chart1.Series[3].Points.Clear();
            chart2.Series[0].Points.Clear();
            chart2.Series[1].Points.Clear();
            dataGridView1.RowCount = ye.Count + 1;
            //chart1.Series[7].Points.Clear();
            chart2.Visible = true;
            for (int i = 0; i < ye.Count; i++)
            {
                dataGridView1.Rows[i + 1].Cells[0].Value = i+1;
                dataGridView1[1, i + 1].Value = Math.Round(x[i], 4);
                dataGridView1[2, i + 1].Value = Math.Round(y[i], 4);
                dataGridView1[3, i + 1].Value = Math.Round(ye[i], 4);
                dataGridView1[4, i + 1].Value = Math.Round(kv[i], 4);
                dataGridView1[5, i + 1].Value = Math.Round(Ysh[i], 4);
                dataGridView1[6, i + 1].Value = Math.Round(Ysh2[i], 4);
                //вывод графиков экспериментальных и аппроксимирующих значений функции
                chart1.Series[0].Points.AddXY(x[i], y[i]);
                chart1.Series[1].Points.AddXY(x[i], ye[i]);
                chart1.Series[2].Points.AddXY(x[i], k1[i]);
                chart1.Series[3].Points.AddXY(x[i], k2[i]);
                chart2.Series[0].Points.AddXY(x[i], Ysh[i]);
                chart2.Series[1].Points.AddXY(x[i], Ysh2[i]);
            }
            TxtSumSkvo.Text = f_o().ToString("F6"); //вывод суммы квадратов отклонений
            double kv_err = Math.Sqrt(f_o()) / n;  //расчет среднекв. ошибки
            TxtMeanErr.Text = kv_err.ToString("F6");
            txty1Y2m.Visible = true; label19.Visible = true;
            txty1X2m.Visible = true; label18.Visible = true;
            txty1Y1m.Visible = true; label17.Visible = true;
            txty1X1m.Visible = true; label16.Visible = true;
            //вывод максимумов
            double xm = p[2];
            double ym = p[0] * Math.Exp(-p[1] * (xm - p[2]) * (xm - p[2])) + p[3] * Math.Exp(-p[4] * (xm - p[5]) * (xm - p[5]));
            txty1X1m.Text = xm.ToString("F4");
            txty1Y1m.Text = ym.ToString("F4");
            xm = p[5];
            ym = p[0] * Math.Exp(-p[1] * (xm - p[2]) * (xm - p[2])) + p[3] * Math.Exp(-p[4] * (xm - p[5]) * (xm - p[5]));
            txty1X2m.Text = xm.ToString("F4");
            txty1Y2m.Text = ym.ToString("F4");


            int max = Ysh.IndexOf(Ysh.Max());

            List<double> numbersListYsh = new List<double>();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                numbersListYsh.Add(Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[5].Value));
            }

            List<double> findPeaksYsh(List<double> numbers)
            {
                List<double> peaksYsh = new List<double>();
                for (int i = 1; i < numbers.Count - 1; i++)
                {
                    if (i > 0 && i < (numbersListYsh.Count - 1))
                    {
                        if (numbers[i] >= numbers[i + 1] && numbers[i] >= numbers[i - 1])
                        {
                            peaksYsh.Add(numbers[i]);
                        }
                    }
                }
                int xv;
                double peak1 = (from number in peaksYsh orderby number descending select number).Distinct().Skip(0).First();
                double peak2 = (from number in peaksYsh orderby number descending select number).Distinct().Skip(1).First();

                xv = numbersListYsh.IndexOf(peak1);
                txtYshX1m.Text = Math.Round(x[xv], 4).ToString();
                txtYshY1m.Text = Math.Round(peak1, 4).ToString();

                xv = numbersListYsh.IndexOf(peak2);
                txtYshX2m.Text = Math.Round(x[xv], 4).ToString();
                txtYshY2m.Text = Math.Round(peak2, 4).ToString();

                xv = numbersListYsh.IndexOf(peaksYsh.Min());
                txtYshX1mi.Text = Math.Round(x[xv], 4).ToString();
                txtYshY1mi.Text = Math.Round(peaksYsh.Min(), 4).ToString();

                return peaksYsh;
            }
            findPeaksYsh(numbersListYsh);

            int min = Ysh.IndexOf(Ysh.Min());

            List<double> numbersListYsh2 = new List<double>();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                numbersListYsh2.Add(Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[6].Value));
            }

            List<double> findPeaksYsh2(List<double> numbers)
            {
                List<double> peaksYsh2 = new List<double>();
                for (int i = 1; i < numbers.Count - 1; i++)
                {
                    if (i > 0 && i < (numbersListYsh2.Count - 1))
                    {
                        if (numbers[i] >= numbers[i + 1] && numbers[i] >= numbers[i - 1])
                        {
                            peaksYsh2.Add(numbers[i]);
                        }
                    }
                }
                int xv;
                //IEnumerable<double> peakwo = peakst.Distinct();
                double peak1 = (from number in peaksYsh2 orderby number descending select number).Distinct().Skip(0).First();
                double peak2 = (from number in peaksYsh2 orderby number descending select number).Distinct().Skip(1).First();

                xv = numbersListYsh2.IndexOf(peak1);
                txtYsh2X1m.Text = Math.Round(x[xv], 4).ToString();
                txtYsh2Y1m.Text = Math.Round(peak1, 4).ToString();

                xv = numbersListYsh2.IndexOf(peak2);
                txtYsh2X2m.Text = Math.Round(x[xv], 4).ToString();
                txtYsh2Y2m.Text = Math.Round(peak2, 4).ToString();

                xv = numbersListYsh2.IndexOf(peaksYsh2.Min());
                txtYsh2X1mi.Text = Math.Round(x[xv], 4).ToString();
                txtYsh2Y1mi.Text = Math.Round(peaksYsh2.Min(), 4).ToString();

                return peaksYsh2;
            }
            findPeaksYsh2(numbersListYsh2);

            T3 = Convert.ToDouble(ye[max]);
            T5 = Convert.ToDouble(ye[min]);

            double max1 = Convert.ToDouble(Ysh.Max());
            double min1 = Convert.ToDouble(Ysh.Min());

            double div = Ysh.Min() / Ysh.Max();

            IPS = max1 / T3;
            IPR = Math.Abs(min1 / T5);

            txtIndexIPS.Text = Math.Round(IPS, 4).ToString();
            txtIndexIPR.Text = Math.Round(IPR, 4).ToString();
            txtIndexYshDiv.Text = Math.Round(Math.Abs(div), 4).ToString();

            double mx1 = Convert.ToDouble(txty1X1m.Text.Trim());
            double mx2 = Convert.ToDouble(txty1X2m.Text.Trim());

            if (mx1 > mx2)
                mafx = mx1;
            else
                mafx = mx2;

            FX30 = mafx + 0.3 * mafx;
            txtIndexT3.Text = Math.Round(FX30, 4).ToString();

            double A1 = k1.Max();
            double A2 = k2.Max();
            double A12div = A1 / A2;
            double A21div = A2 / A1;
            txtIndexA1.Text = Math.Round(A1, 4).ToString();
            txtIndexA2.Text = Math.Round(A2, 4).ToString();
            txtIndexA12Div.Text = Math.Round(A12div, 4).ToString();
            txtIndexA21Div.Text = Math.Round(A21div, 4).ToString();

            double my1 = Convert.ToDouble(txty1Y1m.Text.Trim());
            double my2 = Convert.ToDouble(txty1Y2m.Text.Trim());

            if (my1 > my2)
                mafy = my1;
            else
                mafy = my2;

            FY30 = mafy - 0.3 * mafy;
            txtIndexY30.Text = Math.Round(FY30, 4).ToString();

            txtIndexVDM.Text = Math.Round(mafx, 4).ToString();

            btnStat.Enabled = true;
            BtnStatCalc.Enabled = true;
            button1.Enabled = true;
            //}
            /*catch (Exception ex)
            {
                MessageBox.Show("Сперва нужно загрузить файлы" + Environment.NewLine +
                    ex.Message);
            }*/
        }

        public List<double> F(List<double> x1, List<double> a)      // расчет массива значений аппроксимирующей функции
        {
            List<double> Y = new List<double>();
            Y.Clear();
            for (int i = 0; i < x1.Count; i++)
            {
                double p1 = -a[1] * (x1[i] - a[2]) * (x1[i] - a[2]);
                double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                Y.Insert(i, (a[0] * Math.Exp(p1) + a[3] * Math.Exp(p2)));
            }
            return Y;
        }

        public List<double> Fk1(List<double> x1, List<double> a)      // расчет массива значений аппроксимирующей функции
        {
            List<double> T1 = new List<double>();
            T1.Clear();
            for (int i = 0; i < x1.Count; i++)
            {
                double p1 = -a[1] * (x1[i] - a[2]) * (x1[i] - a[2]);
                T1.Insert(i, (a[0] * Math.Exp(p1)));
            }
            return T1;
        }

        public List<double> Fk1sh(List<double> x1, List<double> a)      // расчет массива значений аппроксимирующей функции
        {
            List<double> T1sh = new List<double>();
            for (int i = 0; i < x1.Count; i++)
            {
                double p1 = -a[1] * ((x1[i] - a[2]) * (x1[i] - a[2]));
                T1sh.Insert(i, (-2) * (a[0] * a[1] * (x1[i] - a[2])) * Math.Exp(p1));
            }
            return T1sh;
        }

        public List<double> Fk2(List<double> x1, List<double> a)      // расчет массива значений аппроксимирующей функции
        {
            List<double> T2 = new List<double>();
            for (int i = 0; i < x1.Count; i++)
            {
                double p2 = -a[4] * (x1[i] - a[5]) * (x1[i] - a[5]);
                T2.Insert(i, (a[3] * Math.Exp(p2)));
            }
            return T2;
        }

        public List<double> Fsh(List<double> x1, List<double> a)      // расчет массива значений аппроксимирующей функции
        {
            List<double> TFsh = new List<double>();
            for (int i = 0; i < x1.Count; i++)
            {
                TFsh.Insert(i, ((-2) * (a[0] * a[1] * (x1[i] - a[2])) *
                    Math.Exp(-a[1] * ((x1[i] - a[2]) * (x1[i] - a[2])))
                    - (2 * a[3] * a[4] * (x1[i] - a[5])
                    * Math.Exp(-a[4] * ((x1[i] - a[5]) * (x1[i] - a[5]))))));
            }
            return TFsh;
        }

        public List<double> Fsh2(List<double> x1, List<double> a)      // расчет массива значений аппроксимирующей функции
        {
            List<double> TFsh2 = new List<double>();
            for (int i = 0; i < x1.Count; i++)
            {
                double p1 = -a[1] * Math.Pow((x1[i] - a[2]), 2);
                double p2 = -a[4] * Math.Pow((x1[i] - a[5]), 2);
                TFsh2.Insert(i, (((-2) * (a[0] * a[1] * Math.Exp(p1))) +
                    (4 * a[0] * Math.Pow(a[1], 2) * Math.Pow((x1[i] - a[2]), 2) * Math.Exp(p1)) -
                    (2 * (a[3] * a[4] * Math.Exp(p2))) +
                    (4 * a[3] * Math.Pow(a[4], 2) * Math.Pow((x1[i] - a[5]), 2) * Math.Exp(p2))));
            }

            return TFsh2;
        }

        private void btnCalc_Click(object sender, EventArgs e)
        {
            sumMean1 = 0;
            sumSKVO1 = 0;
            SKO = 0;
            MSE1 = 0;
            disp = 0;
            VARI = 0;
            n = dataGridView1.Rows.Count - 1;

            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                sumMean1 += Convert.ToDouble(dataGridView1.Rows[i + 1].Cells[3].Value);
            }
            double mean = sumMean1 / n;

            dataGridView4.Rows.Clear();
            dataGridView4.RowCount = n+1;
            dataGridView4.ColumnCount = 4;
            dataGridView4[0, 0].Value = "№";
            dataGridView4[1, 0].Value = "F(x)";
            dataGridView4[2, 0].Value = "M-F(x)";
            dataGridView4[3, 0].Value = "(M-F(x))^2";



            //for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            //{
            //    dataGridView4.Rows.Add();
            //}

            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView4.Rows[i + 1].Cells[0].Value = i + 1;
                dataGridView4.Rows[i + 1].Cells[1].Value = dataGridView1.Rows[i + 1].Cells[3].Value;
                dataGridView4.Rows[i + 1].Cells[2].Value = Math.Round(Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[1].Value) - mean, 4);
                dataGridView4.Rows[i + 1].Cells[3].Value = Math.Round(Math.Pow((Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[1].Value) - mean), 2), 4);
            }

            for (int i = 0; i < dataGridView1.Rows.Count - 1; ++i)
            {
                sumSKVO1 += Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[3].Value);
            }

            SKO = Math.Sqrt( sumSKVO1 / (n - 1));

            disp = Math.Pow(SKO, 2);

            

            //SKO = Math.Sqrt(disp);
            MSE1 = SKO / Math.Sqrt(n);
            VARI = SKO / mean;

            txtMean1.Text = Math.Round(mean, 4).ToString();
            txtSKVO1.Text = Math.Round(sumSKVO1, 4).ToString();
            txtDisp.Text = Math.Round(disp, 4).ToString();
            textSKO.Text = Math.Round(SKO, 4).ToString();
            txtMSE1.Text = Math.Round(MSE1, 4).ToString();
            txtVar.Text = Math.Round(VARI, 4).ToString();
            tb1c = dataGridView4.Rows.Count - 1;
            tb2c = dataGridView2.Rows.Count - 1;

            button8.Enabled = true;

        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                sumMean1 = Convert.ToDouble(txtMean1.Text);
                sumMean2 = Convert.ToDouble(txtMean2.Text);
                MSE1 = Convert.ToDouble(txtMSE1.Text);
                MSE2 = Convert.ToDouble(txtMSE2.Text);
                Student = (sumMean1 - sumMean2) / Math.Sqrt(Math.Pow((MSE1), 2) + Math.Pow((MSE2), 2));
                txtStudent.Text = Math.Round(Student, 4).ToString();
                umnozh = 0;
                umnozh2 = 0;
                tb1c = dataGridView2.Rows.Count - 1;
                tb2c = dataGridView4.Rows.Count - 1;

                sumSKVO1 = Convert.ToDouble(txtSKVO1.Text);
                sumSKVO2 = Convert.ToDouble(txtSKVO2.Text);

                if (tb1c >= tb2c)
                {
                    for (int i = 0; i < tb2c; i++)
                    {
                        umnozh += Convert.ToDouble(dataGridView2.Rows[i + 1].Cells[2].Value) * Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[2].Value);
                    }
                    umnozh2 = Math.Sqrt(Math.Pow((sumSKVO1), 2) * Math.Pow((sumSKVO2), 2));
                    Pirson = umnozh / umnozh2;
                    txtPirson.Text = Math.Round(Pirson, 4).ToString();
                }
                else
                {
                    for (int i = 0; i < tb1c; i++)
                    {
                        umnozh += Convert.ToDouble(dataGridView2.Rows[i + 1].Cells[2].Value) * Convert.ToDouble(dataGridView4.Rows[i + 1].Cells[2].Value);
                    }
                    umnozh2 = Math.Sqrt(Math.Pow((sumSKVO1), 2) * Math.Pow((sumSKVO2), 2));
                    Pirson = umnozh / umnozh2;
                    txtPirson.Text = Math.Round(Pirson, 4).ToString();
                }
                button1.Enabled = true;
            }
            catch (Exception ex)
            {
                txtMean1.BackColor = Color.Yellow;
                txtMean2.BackColor = Color.Yellow;
                txtMSE1.BackColor = Color.Yellow;
                txtMSE2.BackColor = Color.Yellow;
                txtSKVO1.BackColor = Color.Yellow;
                txtSKVO2.BackColor = Color.Yellow;
                MessageBox.Show("Помимо основной таблицы нужно загрузить дополнительную таблицу из базы данных для расчёта критериев" + Environment.NewLine +
                    ex.Message);
                txtMean1.BackColor = Color.WhiteSmoke;
                txtMean2.BackColor = Color.WhiteSmoke;
                txtMSE1.BackColor = Color.WhiteSmoke;
                txtMSE2.BackColor = Color.WhiteSmoke;
                txtSKVO1.BackColor = Color.WhiteSmoke;
                txtSKVO2.BackColor = Color.WhiteSmoke;
            }
        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                Point mousePoint = new Point(e.X, e.Y);

                var XPixel = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.X);
                var YPixel = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Y);


                chart1.ChartAreas[0].CursorX.Interval = 0;
                chart1.ChartAreas[0].CursorY.Interval = 0;

                chart1.ChartAreas[0].CursorX.SetCursorPixelPosition(mousePoint, true);
                chart1.ChartAreas[0].CursorY.SetCursorPixelPosition(mousePoint, true);
                toolTip1.SetToolTip(chart1, "x: " + XPixel.ToString("F3") + " y: " + YPixel.ToString("F3"));
            }
            catch (Exception ex)
            {
                
            }
        }

        private void chart2_MouseMove(object sender, MouseEventArgs e)
        {
            Point mousePoint = new Point(e.X, e.Y);

            var XPixel = chart2.ChartAreas[0].AxisX.PixelPositionToValue(e.X);
            var YPixel = chart2.ChartAreas[0].AxisY.PixelPositionToValue(e.Y);


            chart2.ChartAreas[0].CursorX.Interval = 0;
            chart2.ChartAreas[0].CursorY.Interval = 0;

            chart2.ChartAreas[0].CursorX.SetCursorPixelPosition(mousePoint, true);
            chart2.ChartAreas[0].CursorY.SetCursorPixelPosition(mousePoint, true);
            toolTip1.SetToolTip(chart2, "x: " + XPixel.ToString("F3") + " y: " + YPixel.ToString("F3"));
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            xlApp.Quit();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                int SelectedId = Convert.ToInt32(dataGridView3[0, dataGridView3.CurrentCell.RowIndex].Value);
                ye2.Clear();
                string sql = "select * from Source where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                ye2.Add(Convert.ToDouble(rdr["ye"], CultureInfo.InvariantCulture));
                            }
                        }
                    }
                    c.Close();
                }

                sql = "select * from Stats where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        using (SQLiteDataReader rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                //перевести всё из базы в строковые переменные
                                txtMean2.Text = Convert.ToString(rdr["Mean"]);
                                txtSKVO2.Text = Convert.ToString(rdr["SKVO"]);
                                txtMSE2.Text = Convert.ToString(rdr["MSE"]);
                            }
                        }
                    }
                    c.Close();
                }
                dataGridView2.Rows.Clear();
                dataGridView2.Refresh();

                dataGridView2.RowCount = ye2.Count + 1;
                dataGridView2.ColumnCount = 4;
                dataGridView2[0, 0].Value = "№";
                dataGridView2[1, 0].Value = "F(x)";
                dataGridView2[2, 0].Value = "M-F(x)";
                dataGridView2[3, 0].Value = "(M-F(x))^2";


                for (int i = 0; i < ye2.Count; i++)
                {
                    dataGridView2.Rows[i + 1].Cells[0].Value = i + 1;
                    dataGridView2.Rows[i + 1].Cells[1].Value = ye2[i].ToString("F4");
                    dataGridView2.Rows[i + 1].Cells[2].Value = Convert.ToDouble(dataGridView2.Rows[i + 1].Cells[1].Value) - Convert.ToDouble(txtMean1.Text);
                    dataGridView2.Rows[i + 1].Cells[3].Value = Math.Pow((Convert.ToDouble(dataGridView2.Rows[i + 1].Cells[1].Value) - Convert.ToDouble(txtMean1.Text)), 2);
                }

                button8.Enabled = true;
                //активировать главную вкладку
                tabControl1.SelectTab(TbP5);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Сперва нужно загрузить основную таблицу" + Environment.NewLine +
                    ex.Message);
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                int SelectedId = Convert.ToInt32(dataGridView3[0, dataGridView3.CurrentCell.RowIndex].Value);
                string sql = "delete from Entry where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from Source where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from InitApp where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from Koef where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from AppEx where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from YshEx where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from Ysh2Ex where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from Indexes where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from Stats where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }

                sql = "delete from Addd where id_entry = @SelectedId";
                using (SQLiteConnection c = new SQLiteConnection(ConnectionString))
                {
                    c.Open();
                    using (SQLiteCommand cmd = new SQLiteCommand(sql, c))
                    {
                        cmd.Parameters.AddWithValue("@SelectedId", SelectedId);
                        cmd.ExecuteReader();
                    }
                    c.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Невозможно удалить запись из базы данных." + Environment.NewLine +
                    "Либо она не выделена в таблице, либо произошла иная ошибка" + Environment.NewLine +
                    ex.Message);
            }
            DBLoad();
            
        }

        private void btnStat_Click(object sender, EventArgs e)
        {
            tabControl1.SelectTab(TbP5);
            BtnStatCalc.Enabled = true;
            button1.Enabled = true;
        }

        private void btnLoadST_Click(object sender, EventArgs e)
        {
            DBLoad();
            tabControl1.SelectTab(TbP4);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                Pirson = Convert.ToDouble(txtPirson.Text);
                n = dataGridView2.Rows.Count - 1;
                Student = (Pirson * Math.Sqrt(n - 2)) / Math.Sqrt(1 - Math.Pow((Pirson), 2));
                txtStudent.Text = Math.Round(Student, 4).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Значение для критерия Пирсона отсутствует" + Environment.NewLine +
                    ex.Message);
            }
        }

        public void scan(int nom)  //оптимизация одномерной функции
        {
            Boolean a;
            Double z, z1, d1;
            y2 = F(x, p);
            z = f_o();
            do
            {
                d1 = Math.Abs(h1);
                p[nom] = p[nom] + h1;
                y2 = F(x, p);
                z1 = f_o();
                a = (z1 >= z);
                if (a == true) h1 = -h1 / k;
                z = z1;
            }
            while (a == false && d1 > eps1);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DBLoad();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (chart1.Series[0].Enabled == true)
            {
                chart1.Series[0].Enabled = false;
                chart1.Series[0].IsVisibleInLegend = false;
            }
            else
            {
                chart1.Series[0].Enabled = true;
                chart1.Series[0].IsVisibleInLegend = true;
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (chart1.Series[1].Enabled == true)
            {
                chart1.Series[1].Enabled = false;
                chart1.Series[1].IsVisibleInLegend = false;
            }
            else
            {
                chart1.Series[1].Enabled = true;
                chart1.Series[1].IsVisibleInLegend = true;
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            if (chart1.Series[2].Enabled == true)
            {
                chart1.Series[2].Enabled = false;
                chart1.Series[2].IsVisibleInLegend = false;
            }
            else
            {
                chart1.Series[2].Enabled = true;
                chart1.Series[2].IsVisibleInLegend = true;
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (chart1.Series[3].Enabled == true)
            {
                chart1.Series[3].Enabled = false;
                chart1.Series[3].IsVisibleInLegend = false;
            }
            else
            {
                chart1.Series[3].Enabled = true;
                chart1.Series[3].IsVisibleInLegend = true;
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            if (chart2.Series[0].Enabled == true)
            {
                chart2.Series[0].Enabled = false;
                chart2.Series[0].IsVisibleInLegend = false;
            }
            else
            {
                chart2.Series[0].Enabled = true;
                chart2.Series[0].IsVisibleInLegend = true;
            }
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            if (chart2.Series[1].Enabled == true)
            {
                chart2.Series[1].Enabled = false;
                chart2.Series[1].IsVisibleInLegend = false;
            }
            else
            {
                chart2.Series[1].Enabled = true;
                chart2.Series[1].IsVisibleInLegend = true;
            }
        }

        public double f_o()// вычисление суммы отклонений - целевая функция
        {
            double sum = 0;
            for (int i = 0; i < n; i++)
                sum += (y[i] - y2[i]) * (y[i] - y2[i]);
            return sum;
        }

        public void InitialApproximation()
        {
            extr.Clear();
            extr = Find_max();
            //зададим начальное приближение
            int ind_min = Find_local_min(extr[1], extr[2]);
            p.Insert(0, ((y[extr[1]] + y[extr[1] - 1]) / 2)); p.Insert(1, (y[ind_min] / 2)); p.Insert(2, (x[extr[1]] + x[extr[1] - 1]) / 2);
            p.Insert(3, (y[extr[2]] + y[extr[2] - 1]) / 2); p.Insert(4, (y[ind_min] / 2)); p.Insert(5, (x[extr[2]] + x[extr[2] - 1]) / 2);
            txtA0start.Text = Convert.ToString(p[0]);
            txtB0start.Text = Convert.ToString(p[1]);
            txtC0start.Text = Convert.ToString(p[2]);
            txtA1start.Text = Convert.ToString(p[3]);
            txtB1start.Text = Convert.ToString(p[4]);
            txtC1start.Text = Convert.ToString(p[5]);
        }

        public int Find_local_min(int i1, int i2)  //поиск локального минимума табличной цункции на заданном отрезке
        {
            int min = 0;
            for (int i = i1; i < i2; i++)
            {
                if (y[i] < y[i - 1] && y[i] < y[i + 1])
                {
                    min = i;
                }
            }
            return min;
        }

        public List<int> Find_max() //поиск маскимумов табличной функции+границы отрезка
        {
            List<int> ind_max = new List<int>();
            //ind_max.Clear();
            //ind_max.Insert(0, 0);
            //ind_max.Insert(1, 0);
            //ind_max.Insert(2, 0);
            //ind_max.Insert(3, 0);
            int n1 = 0;
            ind_max.Add(n1);
            for (int i = 1; i < n - 1; i++)
            {
                if (y[i] > y[i - 1] && y[i] > y[i + 1])
                {
                    n1++;
                    ind_max.Insert(n1, i);                    
                }
            }
            ind_max.Insert(3, n - 1);
            return ind_max;
        }

        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Delta < 0)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        chart1.Series[i].IsVisibleInLegend = true;
                    }
                    chart1.ChartAreas[0].AxisX.ScaleView.ZoomReset();
                    chart1.ChartAreas[0].AxisY.ScaleView.ZoomReset();
                }

                if (e.Delta > 0)
                {
                    for (int i = 0; i < 4; i++)
                    {
                        chart1.Series[i].IsVisibleInLegend = false;
                    }

                    double xMin = chart1.ChartAreas[0].AxisX.ScaleView.ViewMinimum;
                    double xMax = chart1.ChartAreas[0].AxisX.ScaleView.ViewMaximum;
                    double yMin = chart1.ChartAreas[0].AxisY.ScaleView.ViewMinimum;
                    double yMax = chart1.ChartAreas[0].AxisY.ScaleView.ViewMaximum;

                    double posXStart = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 2;
                    double posXFinish = chart1.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 2;
                    double posYStart = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 2;
                    double posYFinish = chart1.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 2;

                    chart1.ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);
                    chart1.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch
            {
            }
        }

        private void chart2_MouseWheel(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Delta < 0)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        chart2.Series[i].IsVisibleInLegend = true;
                    }
                    chart2.ChartAreas[0].AxisX.ScaleView.ZoomReset();
                    chart2.ChartAreas[0].AxisY.ScaleView.ZoomReset();
                }

                if (e.Delta > 0)
                {
                    for (int i = 0; i < 2; i++)
                    {
                        chart2.Series[i].IsVisibleInLegend = false;
                    }

                    double xMin = chart2.ChartAreas[0].AxisX.ScaleView.ViewMinimum;
                    double xMax = chart2.ChartAreas[0].AxisX.ScaleView.ViewMaximum;
                    double yMin = chart2.ChartAreas[0].AxisY.ScaleView.ViewMinimum;
                    double yMax = chart2.ChartAreas[0].AxisY.ScaleView.ViewMaximum;

                    double posXStart = chart2.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) - (xMax - xMin) / 2;
                    double posXFinish = chart2.ChartAreas[0].AxisX.PixelPositionToValue(e.Location.X) + (xMax - xMin) / 2;
                    double posYStart = chart2.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / 2;
                    double posYFinish = chart2.ChartAreas[0].AxisY.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / 2;

                    chart2.ChartAreas[0].AxisX.ScaleView.Zoom(posXStart, posXFinish);
                    chart2.ChartAreas[0].AxisY.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch { }
        }
    }
}