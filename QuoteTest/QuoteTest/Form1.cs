using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Reflection;
using System.Diagnostics;
using System.Resources;
using System.Threading;
using System.Globalization;

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
        private System.Drawing.Pen penBlue1 = new System.Drawing.Pen(Color.Blue, 1F);
        private System.Drawing.Pen penGreen1 = new System.Drawing.Pen(Color.Green, 1F);
        private System.Drawing.Pen penRed1 = new System.Drawing.Pen(Color.Red, 1F);
        private System.Drawing.Pen penBlue2 = new System.Drawing.Pen(Color.Blue, 2F);
        private System.Drawing.Pen penGreen2 = new System.Drawing.Pen(Color.Green, 2F);
        private System.Drawing.Pen penRed2 = new System.Drawing.Pen(Color.Red, 2F);
        SolidBrush blueBrush = new SolidBrush(Color.Blue);
        SolidBrush greenBrush = new SolidBrush(Color.Green);
        SolidBrush redBrush = new SolidBrush(Color.Red);

        public static string minutes = "0", sec = "0";
        public static double mhigh = 100000, mlow = 0, mopen = 0, mclose = 0;
        public static string lastminutes = "0";
        public static double lastprice = 0;
        public static bool runK = false, run_cal = false;
        public static int timeInterval = 3;
        private static System.Timers.Timer aTimer30;
        public static double WeightSum = 0;

        List<double> popen = new List<double>();
        List<double> pclose = new List<double>();
        List<double> phigh = new List<double>();
        List<double> plow = new List<double>();
        List<long> pvol = new List<long>();

        List<string> stockCode = new List<string>();
        List<double> stockWeight = new List<double>();
        List<int> stockForce = new List<int>();

        //連線Event OnMktStatusChange (int Status, char* Msg)	與行情發送端連線的狀態,回傳LinkStatus 
        private void axYuantaQuote1_OnMktStatusChange(object sender, AxYuantaQuoteLib._DYuantaQuoteEvents_OnMktStatusChangeEvent e)
        { 
            textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ")+e.msg.ToString();
            if (e.msg.ToString().IndexOf("行情連線結束") >= 0)
            {
                //隔幾秒再連線
                
                textBox_status.Text=DateTime.Now.ToString("HH:mm:ss.fff ")+"行情連線結束，隔5秒重新連線";
                MessageBox.Show(DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線結束，隔5秒重新連線");
                timer1.Enabled = true;
            }
            else if (e.msg.ToString().IndexOf("行情連線失敗") >= 0)
            {
                //隔幾秒再連線
                //可能網路不通
                textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ")+"行情連線失敗，隔5秒重新連線";
                MessageBox.Show(DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線失敗，隔5秒重新連線");
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

            //cal_k(e);
            if (run_cal)
            {
                cal_force(e);
            }         
        }

        public void cal_force(AxYuantaQuoteLib._DYuantaQuoteEvents_OnGetMktAllEvent e)
        {
            int seconds = Int32.Parse(e.matchTime.Substring(4, 2));
            int stockIndex = stockCode.IndexOf(e.symbol);
            if (stockIndex != -1)
            {
                stockForce[stockIndex] += 1;
                //Debug.WriteLine("StockIndex:" + stockIndex);
                //Debug.WriteLine("StockForce:" + stockForce[stockIndex]);
                //Debug.WriteLine("StockWeight:" + stockWeight[stockIndex]);
                //判斷成交價位於買盤或賣盤
                if (e.matchPri == e.bestBuyPri.Split(',')[0])
                {
                    stockForce[stockIndex] += Int32.Parse(e.matchQty);
                    //Debug.WriteLine("StockForce---In:" + Int32.Parse(e.matchQty));
                }
                else if (e.matchPri == e.bestSellPri.Split(',')[0])
                {
                    stockForce[stockIndex] -= Int32.Parse(e.matchQty);
                    //Debug.WriteLine("StockForce---Out:" + Int32.Parse(e.matchQty));
                }
                else
                {
                    Debug.WriteLine("StockForce---Nooooooo");
                }
            }  
        }

        public void cal_result()
        {
            //計算權重
            WeightSum = 0;
            for(int i = 0; i < stockCode.Count; i++)
            {
                WeightSum = 0.01 * stockWeight[i] * stockForce[i];
            }
            Debug.WriteLine("WeightSum:" + WeightSum);
            run_cal = true;
        }

        public void read50stock()
        {
            using (var fs = File.OpenRead(@"C:\50stock.csv"))
            using (var reader = new StreamReader(fs))
            {

                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    stockCode.Add(values[1]);
                    stockWeight.Add(Convert.ToDouble(values[3].TrimEnd('%')));
                    stockForce.Add(0);
                }
            }
        }

        public void load50stock()
        {
            for (int i = 0; i < stockCode.Count; i++)
            {                
                axYuantaQuote1.AddMktReg(stockCode[i], comboBox_UpdateMode.Text.Substring(0, 1));
            }            
        }
        
        public void cal_k(AxYuantaQuoteLib._DYuantaQuoteEvents_OnGetMktAllEvent e)
        {
            double tmp = Convert.ToDouble(e.matchPri);
            //判斷是否第一次記錄
            if (!runK)
            {
                runK = true;
                mhigh = tmp;
                mlow = tmp;
                mopen = tmp;
                mclose = tmp;
                minutes = e.matchTime.Substring(2, 2);
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+"InitializeMinutes=" + minutes);
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+
                    "mhigh=" + mhigh + "_mlow=" + mlow + "_mopen=" + mopen + "_mclose=" + mclose);
                
                draw_k();
            }
            //判斷是否新的一分鐘
            if (String.Compare(e.matchTime.Substring(2, 2),minutes,true)!=0)
            {
                mclose = lastprice;
                /*TODO:開始畫圖*/
                popen.Add(mopen);
                pclose.Add(mclose);
                phigh.Add(mhigh);
                plow.Add(mlow);
                draw_k();
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+"Minutes=" + minutes);
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+
                    "mhigh=" + mhigh+"_mlow="+mlow+"_mopen="+mopen+"_mclose="+mclose);
                //初始化該分鐘的資料
                mhigh = tmp;
                mlow = tmp;
                mopen = tmp;
                mclose = tmp;
                minutes = e.matchTime.Substring(2, 2);
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+"NewMinutes=" + minutes);
                //MessageBox.Show("SetMktConnection失敗：" + mhigh);
            }
            //比較高低價
            if (tmp > mhigh)
            {
                mhigh = tmp;
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+"NewHigh=" + mhigh);
            }
            if (tmp < mlow)
            {
                mlow = tmp;
                Debug.WriteLine(DateTime.Now.ToString("h:mm:ss ")+"NewLow=" + mlow);
            }
            mclose = tmp;
            lastprice = tmp;
        }

        public void draw_k()
        {
            g = pictureBox1.CreateGraphics();
            Size cs = pictureBox1.ClientSize;
            int scale = 1, shift = 0;
            long trademax = 0;
            double max = 0, min = 10000;
            int kx = cs.Width - 10;

            g.Clear(Color.Black);
            //計算能畫幾條k棒
            int amount = (cs.Width) / 10 / scale + 1;
            //計算最高最低價與最高成交量
            for (int i = shift; i < pvol.Count; i++)
            {
                if (pvol[i] > trademax)
                {
                    trademax = pvol[i];
                    //Debug.WriteLine("UpdateNewTrade=" + trademax);
                }
            }

            for (int i = shift; i < phigh.Count; i++)
            {
                if (phigh[i] > max)
                {
                    max = phigh[i];
                }                    
            }

            for (int i = shift; i < plow.Count; i++)
            {
                if (plow[i] < min)
                {
                    min = plow[i];
                }                    
            }
            //繪製目前價格
            for(int i = shift; i < popen.Count; i++)
            {
                int kxWidth = 5;
                kx = kx - kxWidth * scale;
                double highy = (cs.Height / (max - min)) * (max - phigh[i]);
                double lowy = (cs.Height / (max - min)) * (max - plow[i]);
                double openy = (cs.Height / (max - min)) * (max - popen[i]);
                double closey = (cs.Height / (max - min)) * (max - pclose[i]);
                Debug.WriteLine("yhigh" + highy);
                Debug.WriteLine("ylow" + lowy);
                Debug.WriteLine("yopen" + openy);
                Debug.WriteLine("yclose" + closey);
                // Create rectangle.
                // Draw rectangle to screen.
                Debug.WriteLine("kx" + kx);
                
                if (popen[i] > pclose[i])//跌
                {
                    Rectangle rect = new Rectangle(kx, (int)openy, kxWidth*2, (int)(closey - openy));
                    g.DrawLine(penGreen2, kx+kxWidth, (float)highy, kx + kxWidth, (float)lowy);
                    g.DrawRectangle(penGreen1, rect);
                    g.FillRectangle(greenBrush, rect);
                }
                else//漲
                {
                    Rectangle rect = new Rectangle(kx, (int)closey, kxWidth*2, (int)(openy - closey));
                    g.DrawLine(penRed2, kx + kxWidth, (float)highy, kx + kxWidth, (float)lowy);
                    g.DrawRectangle(penRed1, rect);
                    g.FillRectangle(redBrush, rect);
                }
                kx = kx - kxWidth * 2 - 2;
            }
        }

        private void button_login_Click(object sender, EventArgs e)
        {
            LoginFn();
            read50stock();
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
                load50stock();
            }
            catch (Exception ex)
            {
                MessageBox.Show("DelMktReg失敗：" + ex.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
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
        
        private void btn_Load_Click(object sender, EventArgs e)
        {
            aTimer30 = new System.Timers.Timer(30000);            
            aTimer30.Elapsed += OnTimedEvent;
            aTimer30.AutoReset = true;
            aTimer30.Enabled = true;
            run_cal = true;
            load50stock();
        }

        private void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {
            Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);
            run_cal = false;
            cal_result();
        }

        private void btn_Writefile_Click(object sender, EventArgs e)
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
