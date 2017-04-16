using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QuoteTest
{
    public partial class Form1 : Form
    {
        public DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }

        private System.Drawing.Graphics g;
        private System.Drawing.Pen pen1 = new System.Drawing.Pen(Color.Blue, 2F);

        //連線Event OnMktStatusChange (int Status, char* Msg)	與行情發送端連線的狀態,回傳LinkStatus 
        private void axYuantaQuote1_OnMktStatusChange(object sender, AxYuantaQuoteLib._DYuantaQuoteEvents_OnMktStatusChangeEvent e)
        { 
            textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ")+e.msg.ToString();
            if (e.msg.ToString().IndexOf("行情連線結束") >= 0)
            {
                //隔幾秒再連線
                
                textBox_status.Text=DateTime.Now.ToString("HH:mm:ss.fff ")+"行情連線結束，隔5秒重新連線";
                timer1.Enabled = true;
            }
            else if (e.msg.ToString().IndexOf("行情連線失敗") >= 0)
            {
                //隔幾秒再連線
                //可能網路不通
                textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ")+"行情連線失敗，隔5秒重新連線";
                timer1.Enabled = true;
            }
        }

        private void axYuantaQuote1_OnGetMktAll(object sender, AxYuantaQuoteLib._DYuantaQuoteEvents_OnGetMktAllEvent e)
        {
            //DataGrid.Rows.Add(e.symbol, e.refPri, e.openPri, e.highPri, e.lowPri, e.upPri, e.dnPri, e.matchTime, e.matchPri, e.matchQty, e.tolMatchQty);
            DataRow DR = this.dt.Rows.Find(e.symbol);
            if (DR != null)
            {
                DR["參考價"] = e.refPri;
                DR["開盤價"] = e.openPri;
                DR["最高價"] = e.highPri;
                DR["最低價"] = e.lowPri;
                DR["漲停價"] = e.upPri;
                DR["跌停價"] = e.dnPri;
                DR["成交時間"] = e.matchTime != "" ? (string.Format("{0}:{1}:{2}.{3}", e.matchTime.Substring(0, 2), e.matchTime.Substring(2, 2), e.matchTime.Substring(4, 2), e.matchTime.Substring(6, 3))) : "";
                DR["成交價位"] = e.matchPri;
                DR["成交數量"] = e.matchQty;
                DR["總成交量"] = e.tolMatchQty;
                DR["買五"] = e.bestBuyPri +e.bestBuyQty ;
                DR["賣五"] = e.bestSellPri+e.bestSellQty;

                DR.EndEdit();
                DR.AcceptChanges();
            }
            else
            {
                DR = this.dt.NewRow();
                DR["商品代碼"] = e.symbol;
                DR["參考價"] = e.refPri;
                DR["開盤價"] = e.openPri;
                DR["最高價"] = e.highPri;
                DR["最低價"] = e.lowPri;
                DR["漲停價"] = e.upPri;
                DR["跌停價"] = e.dnPri;
                DR["成交時間"] = e.matchTime!="" ? (string.Format("{0}:{1}:{2}.{3}", e.matchTime.Substring(0, 2), e.matchTime.Substring(2, 2), e.matchTime.Substring(4, 2), e.matchTime.Substring(6, 3)) ): "";
                DR["成交價位"] = e.matchPri;
                DR["成交數量"] = e.matchQty;
                DR["總成交量"] = e.tolMatchQty;
                DR["買五"] = e.bestBuyPri + e.bestBuyQty;
                DR["賣五"] = e.bestSellPri + e.bestSellQty;
                this.dt.Rows.Add(DR);
            }

            ListShow(String.Format("{0} 買五：{1}-{2}", e.symbol, e.bestBuyPri, e.bestBuyQty));
            ListShow(String.Format("{0} 賣五：{1}-{2}", e.symbol, e.bestSellPri, e.bestSellQty));
        }

        private void button_login_Click(object sender, EventArgs e)
        {
            LoginFn();
        }

        private void LoginFn()
        {
            try
            {
                axYuantaQuote1.SetMktLogon(textBox_id.Text.Trim(), textBox_pass.Text.Trim(), textBox_ip.Text.Trim(), textBox_port.Text.Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show("SetMktConnection失敗：" + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                int RegErrCode = axYuantaQuote1.AddMktReg(textBox_sym.Text.Trim(), comboBox_UpdateMode.Text.Substring(0, 1));
                textBox_status2.Text = DateTime.Now.ToString("HH:mm:ss.fff ")+RegErrCode.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("DelMktReg失敗：" + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            g = pictureBox1.CreateGraphics();
            g.DrawEllipse(pen1, 50, 50, 100, 150);
            try
            {
                int RegErrCode = axYuantaQuote1.DelMktReg(textBox_sym.Text.Trim());
                textBox_status2.Text = RegErrCode.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("DelMktReg失敗：" + ex.Message);
            }
        }

        private void axYuantaQuote1_OnRegError(object sender, AxYuantaQuoteLib._DYuantaQuoteEvents_OnRegErrorEvent e)
        {
            textBox_status2.Text = e.errCode.ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            this.dt = new DataTable("QuoteTable");
            DataColumn Col0 = new DataColumn("商品代碼", System.Type.GetType("System.String"));
             DataColumn Col2 = new DataColumn("參考價", System.Type.GetType("System.String"));
            DataColumn Col3 = new DataColumn("開盤價", System.Type.GetType("System.String"));
            DataColumn Col4 = new DataColumn("最高價", System.Type.GetType("System.String"));
            DataColumn Col5 = new DataColumn("最低價", System.Type.GetType("System.String"));
            DataColumn Col6 = new DataColumn("漲停價", System.Type.GetType("System.String"));
            DataColumn Col7 = new DataColumn("跌停價", System.Type.GetType("System.String"));
            DataColumn Col8 = new DataColumn("成交時間", System.Type.GetType("System.String"));
            DataColumn Col9 = new DataColumn("成交價位", System.Type.GetType("System.String"));
            DataColumn Col10 = new DataColumn("成交數量", System.Type.GetType("System.String"));
            DataColumn Col11 = new DataColumn("總成交量", System.Type.GetType("System.String"));
            DataColumn Col12 = new DataColumn("買五", System.Type.GetType("System.String"));
            DataColumn Col13 = new DataColumn("賣五", System.Type.GetType("System.String"));
            this.dt.Columns.Add(Col0);
             this.dt.Columns.Add(Col2);
            this.dt.Columns.Add(Col3);
            this.dt.Columns.Add(Col4);
            this.dt.Columns.Add(Col5);
            this.dt.Columns.Add(Col6);
            this.dt.Columns.Add(Col7);
            this.dt.Columns.Add(Col8);
            this.dt.Columns.Add(Col9);
            this.dt.Columns.Add(Col10);
            this.dt.Columns.Add(Col11);
            this.dt.Columns.Add(Col12);
            this.dt.Columns.Add(Col13);
            this.dt.PrimaryKey = new DataColumn[] { this.dt.Columns["商品代碼"] };
            bindingSource1.DataSource = this.dt;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            LoginFn();
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            
        }

        private void statusStrip1_Click(object sender, EventArgs e)
        {
            panel1.Hide();
        }

        private void statusStrip1_DoubleClick(object sender, EventArgs e)
        {
            panel1.Show();
        }

        private void DataGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            try
            {
                if (e.Value.ToString() != "")
                {
                    if (this.DataGrid.Columns["UpPri"].Index == e.ColumnIndex && e.RowIndex >= 0)
                    {
                        e.CellStyle.BackColor = Color.Red;
                    }

                    if (this.DataGrid.Columns["DnPri"].Index == e.ColumnIndex && e.RowIndex >= 0)
                    {
                        e.CellStyle.BackColor = Color.Green;
                    }

                    if ((this.DataGrid.Columns["OpenPri"].Index == e.ColumnIndex && e.RowIndex >= 0)||(this.DataGrid.Columns["HighPri"].Index == e.ColumnIndex && e.RowIndex >= 0)|| (this.DataGrid.Columns["LowPri"].Index == e.ColumnIndex && e.RowIndex >= 0)|| (this.DataGrid.Columns["MatchPri"].Index == e.ColumnIndex && e.RowIndex >= 0))
                    {
                        if (string.Compare(DataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), DataGrid.Rows[e.RowIndex].Cells[5].Value.ToString()) == 0)
                            e.CellStyle.BackColor = Color.Red;
                        else if (string.Compare(DataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), DataGrid.Rows[e.RowIndex].Cells[6].Value.ToString()) == 0)
                            e.CellStyle.BackColor = Color.Green;
                        else if (string.Compare(DataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), DataGrid.Rows[e.RowIndex].Cells[1].Value.ToString()) > 0)
                            e.CellStyle.ForeColor = Color.Red;
                        else if (string.Compare(DataGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), DataGrid.Rows[e.RowIndex].Cells[1].Value.ToString()) < 0)
                            e.CellStyle.ForeColor = Color.Lime;
                        else
                            e.CellStyle.ForeColor = Color.White;
                    }


                    
                }

            }
            catch (Exception ex)
            { }
        }

        private void textBox_id_Click(object sender, EventArgs e)
        {
            textBox_id.SelectAll();
        }

        private void textBox_sym_Click(object sender, EventArgs e)
        {
            textBox_sym.SelectAll();
        }

        private delegate void InvokeFunction(string msg);
        public void ListShow(string str_log)
        {
            string StrLog = string.Format("{0}  [{1}] ", DateTime.Now.ToString("HH:mm:ss.fff "), str_log);
            listBox1.BeginInvoke(new InvokeFunction(ShowMsg), new object[] { StrLog });
        }

        public void ShowMsg(string logstr)
        {
            this.Invoke((System.Windows.Forms.MethodInvoker)delegate
            {
                if (listBox1.Items.Count > 1000)
                {
                    listBox1.Items.Clear();
                    listBox1.Items.Insert(0, string.Format("{0}  [{1}]", DateTime.Now.ToString("HH:mm:ss.fff "), "清除舊資料"));
                }
                listBox1.Items.Insert(0, logstr);
            });
        }

        private void DataGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void folderBrowserDialog1_HelpRequest(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox_UpdateMode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {

        }
    }
}
