using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NewRegression
{
    public partial class Stat : Form
    {
        Form1 frmMain = new Form1();
        public Stat()
        {
            InitializeComponent();
            
            dataGridView1.RowCount = 1;    //задаем кол-во строк и столбцов каждой таблицы на форме
            dataGridView1.ColumnCount = 14;

            dataGridView2.RowCount = 1;    //задаем кол-во строк и столбцов каждой таблицы на форме
            dataGridView2.ColumnCount = 4;

            
        }            
    }
}
