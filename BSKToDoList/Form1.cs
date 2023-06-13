using System;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace BSKToDoList
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        OleDbConnection con;
        OleDbDataAdapter da;
        OleDbCommand cmd;
        DataSet ds;

        void griddoldur()
        {
            con = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\ToDo\\bsk.mdb");
            da = new OleDbDataAdapter("Select * from todo", con);
            ds = new DataSet();
            con.Open();
            da.Fill(ds, "todo");
            dataGridView3.DataSource = ds.Tables["todo"];
            dataGridView4.DataSource = ds.Tables["todo"];
            con.Close();



        }

        void kayitdoldur()
        {

            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.Cells[0].Value != null)
                {
                    if ((Boolean)row.Cells[0].Value == true)
                    {
                        dataGridView2.Rows.Add(row.Cells[0].Value, row.Cells[1].Value, row.Cells[2].Value.ToString(), row.Cells[3].Value.ToString(), row.Cells[4].Value.ToString(), row.Cells[5].Value.ToString(), row.Cells[6].Value.ToString());

                    }
                    else
                        dataGridView1.Rows.Add(row.Cells[0].Value, row.Cells[1].Value, row.Cells[2].Value.ToString(), row.Cells[3].Value.ToString(), row.Cells[4].Value.ToString(), row.Cells[5].Value.ToString(), row.Cells[6].Value.ToString());
                }
            }
        }
        void delete()
        {
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "delete from todo where ToDo_ID=" + textBox2.Text + "";
            cmd.ExecuteNonQuery();
            con.Close();
            griddoldur();
            textBox1.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();
            griddoldur();
            kayitdoldur();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Column14.FillWeight = 100;
            Column3.FillWeight = 100;
            Column4.FillWeight = 100;
            Column5.FillWeight = 100;
            Column6.FillWeight = 100;
            Column7.FillWeight = 100;
            Column9.FillWeight = 100;
            Column10.FillWeight = 100;
            Column11.FillWeight = 100;
            Column12.FillWeight = 100;
            statusDataGridViewCheckBoxColumn.FillWeight = 100;
            istipiDataGridViewTextBoxColumn.FillWeight = 100;
            urgDataGridViewTextBoxColumn.FillWeight = 100;
            periodDataGridViewTextBoxColumn.FillWeight = 100;
            impDataGridViewTextBoxColumn.FillWeight = 100;


            WindowState = FormWindowState.Maximized;

            griddoldur();

            // TODO: This line of code loads data into the 'bskDataSet1.period' table. You can move, or remove it, as needed.
            this.periodTableAdapter.Fill(this.bskDataSet1.period);
            // TODO: This line of code loads data into the 'bskDataSet1.urg' table. You can move, or remove it, as needed.
            this.urgTableAdapter.Fill(this.bskDataSet1.urg);
            // TODO: This line of code loads data into the 'bskDataSet1.imp' table. You can move, or remove it, as needed.
            this.impTableAdapter.Fill(this.bskDataSet1.imp);
            // TODO: This line of code loads data into the 'bskDataSet1.is_tipi' table. You can move, or remove it, as needed.
            this.is_tipiTableAdapter.Fill(this.bskDataSet1.is_tipi);
            // TODO: This line of code loads data into the 'bskDataSet.todo' table. You can move, or remove it, as needed.
            this.todoTableAdapter.Fill(this.bskDataSet.todo);

            dataGridView1.RowHeadersVisible = false;
            dataGridView2.RowHeadersVisible = false;
            dataGridView3.RowHeadersVisible = false;
            textBox1.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "INSERT INTO todo (status,topic,istipi,imp,urg,period) VALUES (@status,@topic,@istipi,@imp,@urg, @period)";
            cmd.Parameters.AddWithValue("@status", false);
            cmd.Parameters.AddWithValue("@topic", textBox1.Text);
            cmd.Parameters.AddWithValue("@istipi", comboBox1.GetItemText(comboBox1.SelectedItem));
            cmd.Parameters.AddWithValue("@imp", comboBox2.GetItemText(comboBox2.SelectedItem));
            cmd.Parameters.AddWithValue("@urg", comboBox3.GetItemText(comboBox3.SelectedItem));
            cmd.Parameters.AddWithValue("@period", comboBox4.GetItemText(comboBox4.SelectedItem));

            if (cmd.ExecuteNonQuery() > 0)
            {
                MessageBox.Show("Record inserted");
            }
            else
            {
                MessageBox.Show("Record failed");
            }
            con.Close();

            textBox1.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();

            griddoldur();
            kayitdoldur();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "delete from todo where ToDo_ID=" + textBox2.Text + "";
            cmd.ExecuteNonQuery();
            con.Close();
            griddoldur();
            textBox1.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();

            griddoldur();
            kayitdoldur();

        }

        private void button3_Click(object sender, EventArgs e)
        {

            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "update todo set topic='" + textBox1.Text + "',istipi='" + comboBox1.SelectedValue + "',imp='" + comboBox2.SelectedValue + "',urg='" + comboBox3.SelectedValue + "',period='" + comboBox4.SelectedValue + "' where ToDo_ID=" + textBox2.Text + "";
            cmd.ExecuteNonQuery();
            con.Close();
            griddoldur();

            textBox1.Clear();
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();

            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();

            griddoldur();
            kayitdoldur();
        }

        private void pictureBox2_Click_1(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0) { GetExcel1(); }
            if (tabControl1.SelectedIndex == 1) { GetExcel2(); }
            if (tabControl1.SelectedIndex == 2) { GetExcel3(); }
        }
        public void GetExcel1()
        {
            try
            {
                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("This process may take a long time depending on the data density. Do you want to continue?", "EXCEL'E AKTARMA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
                    uyg.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
                        myRange.Value2 = dataGridView1.Columns[i].HeaderText;
                    }
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        for (int j = 0; j < dataGridView1.Rows.Count; j++)
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                            myRange.Value2 = dataGridView1[i, j].Value;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("THE OPERATION HAS BEEN CANCELED.", "Operation Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("YOU CLOSED THE EXCEL WINDOW BEFORE THE PROCESS IS COMPLETED.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void GetExcel2()
        {
            try
            {
                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("This process may take a long time depending on the data density. Do you want to continue?", "EXCEL'E AKTARMA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
                    uyg.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
                    for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
                        myRange.Value2 = dataGridView2.Columns[i].HeaderText;
                    }

                    for (int i = 0; i < dataGridView2.Columns.Count; i++)
                    {
                        for (int j = 0; j < dataGridView2.Rows.Count; j++)
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                            myRange.Value2 = dataGridView2[i, j].Value;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("THE OPERATION HAS BEEN CANCELED.", "Operation Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("YOU CLOSED THE EXCEL WINDOW BEFORE THE PROCESS IS COMPLETED.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void GetExcel3()
        {
            try
            {
                DialogResult dialog = new DialogResult();
                dialog = MessageBox.Show("This process may take a long time depending on the data density. Do you want to continue?", "EXCEL'E AKTARMA", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    Microsoft.Office.Interop.Excel.Application uyg = new Microsoft.Office.Interop.Excel.Application();
                    uyg.Visible = true;
                    Microsoft.Office.Interop.Excel.Workbook kitap = uyg.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet sheet1 = (Microsoft.Office.Interop.Excel.Worksheet)kitap.Sheets[1];
                    for (int i = 0; i < dataGridView3.Columns.Count; i++)
                    {
                        Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[1, i + 1];
                        myRange.Value2 = dataGridView3.Columns[i].HeaderText;
                    }

                    for (int i = 0; i < dataGridView3.Columns.Count; i++)
                    {
                        for (int j = 0; j < dataGridView3.Rows.Count; j++)
                        {
                            Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
                            myRange.Value2 = dataGridView3[i, j].Value;
                        }
                    }
                }
                else
                {
                    MessageBox.Show("THE OPERATION HAS BEEN CANCELED.", "Operation Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception)
            {

                MessageBox.Show("YOU CLOSED THE EXCEL WINDOW BEFORE THE PROCESS IS COMPLETED.", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }

        private void pictureBox4_Click_1(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0) { GetRecord1(); }
            if (tabControl1.SelectedIndex == 1) { GetRecord2(); }
        }
        public void GetRecord1()
        {
            int kayitSayisi;
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "SELECT COUNT(*) FROM todo where (status=0)";
            kayitSayisi = (int)cmd.ExecuteScalar();
            con.Close();
            MessageBox.Show("Total record: " + kayitSayisi.ToString());
        }
        public void GetRecord2()
        {
            int kayitSayisi;
            cmd = new OleDbCommand();
            con.Open();
            cmd.Connection = con;
            cmd.CommandText = "SELECT COUNT(*) FROM todo where (status=true)";
            kayitSayisi = (int)cmd.ExecuteScalar();
            con.Close();
            MessageBox.Show("Total record: " + kayitSayisi.ToString());
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            kayitdoldur();
        }
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string deger = dataGridView1.SelectedCells[1].Value.ToString();
            textBox2.Text = deger.ToString();

            string deger1 = dataGridView1.SelectedCells[2].Value.ToString();
            textBox1.Text = deger1.ToString();

            string deger2 = dataGridView1.SelectedCells[3].Value.ToString();
            comboBox1.Text = deger2.ToString();

            string deger3 = dataGridView1.SelectedCells[4].Value.ToString();
            comboBox3.Text = deger3.ToString();

            string deger4 = dataGridView1.SelectedCells[5].Value.ToString();
            comboBox4.Text = deger4.ToString();

            string deger5 = dataGridView1.SelectedCells[6].Value.ToString();
            comboBox2.Text = deger5.ToString();

        }
        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string deger = dataGridView2.SelectedCells[1].Value.ToString();
            textBox2.Text = deger.ToString();

            string deger1 = dataGridView2.SelectedCells[2].Value.ToString();
            textBox1.Text = deger1.ToString();

            string deger2 = dataGridView2.SelectedCells[3].Value.ToString();
            comboBox1.Text = deger2.ToString();

            string deger3 = dataGridView2.SelectedCells[4].Value.ToString();
            comboBox3.Text = deger3.ToString();

            string deger4 = dataGridView2.SelectedCells[5].Value.ToString();
            comboBox4.Text = deger4.ToString();

            string deger5 = dataGridView2.SelectedCells[6].Value.ToString();
            comboBox2.Text = deger5.ToString();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            if (dataGridView1.CurrentCell.ColumnIndex.Equals(0))
            {
                foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        cmd = new OleDbCommand();
                        con.Open();
                        cmd.Connection = con;
                        int kayitno = Convert.ToInt32(dataGridView1.CurrentRow.Cells[1].Value);
                        cmd.CommandText = "update todo set status=" + true + " where ToDo_ID=" + kayitno + "";
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }
                }

                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();

                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();

                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
                griddoldur();
                foreach (DataGridViewRow row2 in dataGridView3.Rows)
                {
                    if (row2.Cells[0].Value != null)
                    {
                        if ((Boolean)row2.Cells[0].Value == true)
                        {
                            dataGridView2.Rows.Add(row2.Cells[0].Value, row2.Cells[1].Value, row2.Cells[2].Value.ToString(), row2.Cells[3].Value.ToString(), row2.Cells[4].Value.ToString(), row2.Cells[5].Value.ToString(), row2.Cells[6].Value.ToString());

                        }
                        else
                            dataGridView1.Rows.Add(row2.Cells[0].Value, row2.Cells[1].Value, row2.Cells[2].Value.ToString(), row2.Cells[3].Value.ToString(), row2.Cells[4].Value.ToString(), row2.Cells[5].Value.ToString(), row2.Cells[6].Value.ToString());
                    }
                }
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentCell.ColumnIndex.Equals(0))
            {
                foreach (DataGridViewRow row in dataGridView2.SelectedRows)
                {
                    if (row.Cells[0].Value != null)
                    {
                        cmd = new OleDbCommand();
                        con.Open();
                        cmd.Connection = con;
                        int kayitno = Convert.ToInt32(dataGridView2.CurrentRow.Cells[1].Value);
                        cmd.CommandText = "update todo set status=" + false + " where ToDo_ID=" + kayitno + "";
                        cmd.ExecuteNonQuery();
                        con.Close();
                    }

                }
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();

                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();

                dataGridView3.DataSource = null;
                dataGridView3.Rows.Clear();
                griddoldur();
                foreach (DataGridViewRow row2 in dataGridView3.Rows)
                {
                    if (row2.Cells[0].Value != null)
                    {
                        if ((Boolean)row2.Cells[0].Value == true)
                        {
                            dataGridView2.Rows.Add(row2.Cells[0].Value, row2.Cells[1].Value, row2.Cells[2].Value.ToString(), row2.Cells[3].Value.ToString(), row2.Cells[4].Value.ToString(), row2.Cells[5].Value.ToString(), row2.Cells[6].Value.ToString());

                        }
                        else
                            dataGridView1.Rows.Add(row2.Cells[0].Value, row2.Cells[1].Value, row2.Cells[2].Value.ToString(), row2.Cells[3].Value.ToString(), row2.Cells[4].Value.ToString(), row2.Cells[5].Value.ToString(), row2.Cells[6].Value.ToString());
                    }
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();

                DataView dv = ds.Tables["todo"].DefaultView;
                dv.RowFilter = "topic LIKE '%" + textBox3.Text + "%'";
                dataGridView4.DataSource = dv;


                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    if ((Boolean)row.Cells[0].Value == false)
                    {
                        object[] rowData = new object[row.Cells.Count];
                        for (int i = 0; i < rowData.Length; ++i)
                        {
                            rowData[i] = row.Cells[i].Value;
                        }
                        this.dataGridView1.Rows.Add(rowData);
                    }
                }

            }

            if (tabControl1.SelectedIndex == 1)
            {
                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();

                DataView dv = ds.Tables["todo"].DefaultView;
                dv.RowFilter = "topic LIKE '%" + textBox3.Text + "%'";
                dataGridView4.DataSource = dv;


                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    if ((Boolean)row.Cells[0].Value == true)
                    {
                        object[] rowData = new object[row.Cells.Count];
                        for (int i = 0; i < rowData.Length; ++i)
                        {
                            rowData[i] = row.Cells[i].Value;
                        }
                        this.dataGridView2.Rows.Add(rowData);
                    }
                }

            }
        }

        private void dataGridView3_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            string deger = dataGridView3.SelectedCells[1].Value.ToString();
            textBox2.Text = deger.ToString();

            string deger1 = dataGridView3.SelectedCells[2].Value.ToString();
            textBox1.Text = deger1.ToString();

            string deger2 = dataGridView3.SelectedCells[3].Value.ToString();
            comboBox1.Text = deger2.ToString();

            string deger3 = dataGridView3.SelectedCells[4].Value.ToString();
            comboBox3.Text = deger3.ToString();

            string deger4 = dataGridView3.SelectedCells[5].Value.ToString();
            comboBox4.Text = deger4.ToString();

            string deger5 = dataGridView3.SelectedCells[6].Value.ToString();
            comboBox2.Text = deger5.ToString();
        }
    }

    }

