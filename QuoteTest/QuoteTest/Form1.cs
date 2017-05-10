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
        //宣告k線畫筆
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
        //計算K棒變數宣告
        public static string minutes = "0", sec = "0";
        public static double mhigh = 100000, mlow = 0, mopen = 0, mclose = 0;
        public static string lastminutes = "0";
        public static double lastprice = 0;
        public static bool runK = false;
        //計算權重變數宣告
        public static int timeInterval = 2;
        public static int btimeInterval = 1;
        public static int tradetimeInterval = 10;
        public static int calInterval = 30;
        private static System.Timers.Timer aTimer30;
        private static System.Timers.Timer bTimer;
        private static System.Timers.Timer tradeTimer;
        public static double WeightSum = 0;
        public static double txfWeight = 0;
        //宣告K線計算開高低收成交
        List<double> popen = new List<double>();
        List<double> pclose = new List<double>();
        List<double> phigh = new List<double>();
        List<double> plow = new List<double>();
        List<long> pvol = new List<long>();
        //宣告力道計算變數
        List<string> stockCode = new List<string>();//商品代碼
        List<double> stockWeight = new List<double>();//商品權重
        List<int> stockForce = new List<int>();//商品力道
        List<double> stockLastPrice = new List<double>();//商品上次價格
        List<int> stockState = new List<int>();//使否計算力道
        List<int> stockLastAmount = new List<int>();//商品上次成交量
        List<List<int>> stockTimeForce = new List<List<int>>();//商品每秒力道
                                                               //交易策略變數宣告
                                                               //TextBox:(1)台指期力道[Display,0](2)權值股力道[Display,0](3)買賣建議方向[Display](4)目前手中持有台指期口數[Display]
                                                               //          (5)目前盈餘[Display,0](6)初始投資金額[1~9999999,100000](7)交易手續費[0~9999999,0](8)停利報酬[0~9999999,5](9)停損報酬[0~9999999,10]
                                                               //          (10)策略一:買賣力道差距參數[0~500,20](11)策略二:買賣力道差距倍數參數[0~50,5]
                                                               //          (12)策略一開關["開"關](13)策略二開關["開"關](14)力道計算秒數[0~59,30](15)載入(16)手動買賣(17)空button*2
                                                               //(18)交易間隔秒數[0~59,5]
        public static int tradeDIR = 0;//買(1)or賣(-1)or不買賣(0)台指期(textbox)
        public static int TXFamount = 0;//目前持有台指期口數(textbox)
        public static double totalAward = 0;//目前交易盈餘(textbox)
        List<double> TXFpriceP = new List<double>();//目前持有台指期多單價格表
        List<double> TXFpriceN = new List<double>();//目前持有台指期空單價格表
        public static double initMoney = 100000;//初始投資金額(textbox)
        public static int strategy1_BuySellLength = 20;//權值股台指期買賣力道差距參數(textbox)
        public static bool strategy1_autoTrade = true;//策略一自動交易(小button)
        public static int strategy2_BuySellmultiply = 5;//權值股台指期買賣力道倍數參數(textbox)
        public static bool strategy2_autoTrade = true;//策略二自動交易(小button)
        public static int fee = 1;//交易手續費(textbox)
        public static int award = 2;//想要停利的報酬(textbox)
        public static int penalty = 2;//想要停損的金額點數(textbox)
        public static double win = 0;//賺錢次數
        public static double lose = 0;//賠錢次數

        //連線Event OnMktStatusChange (int Status, char* Msg)	與行情發送端連線的狀態,回傳LinkStatus 
        private void axYuantaQuote1_OnMktStatusChange(object sender, AxYuantaQuoteLib._DYuantaQuoteEvents_OnMktStatusChangeEvent e)
        {
            textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + e.msg.ToString();
            if (e.msg.ToString().IndexOf("行情連線結束") >= 0)
            {
                //隔幾秒再連線
                textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線結束，隔5秒重新連線";
                MessageBox.Show(DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線結束，隔5秒重新連線");
                timer1.Enabled = true;
            }
            else if (e.msg.ToString().IndexOf("行情連線失敗") >= 0)
            {
                //隔幾秒再連線
                //可能網路不通
                textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線失敗，隔5秒重新連線";
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
                DR["買五"] = e.bestBuyPri + e.bestBuyQty;
                DR["賣五"] = e.bestSellPri + e.bestSellQty;

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
                DR["成交時間"] = e.matchTime != "" ? (string.Format("{0}:{1}:{2}.{3}", e.matchTime.Substring(0, 2), e.matchTime.Substring(2, 2), e.matchTime.Substring(4, 2), e.matchTime.Substring(6, 3))) : "";
                DR["成交價位"] = e.matchPri;
                DR["成交數量"] = e.matchQty;
                DR["總成交量"] = e.tolMatchQty;
                DR["買賣力道"] = stockForce[stockCode.IndexOf(e.symbol)];
                DR["權重"] = stockWeight[stockCode.IndexOf(e.symbol)];
                DR["買五"] = e.bestBuyPri + e.bestBuyQty;
                DR["賣五"] = e.bestSellPri + e.bestSellQty;
                this.dt.Rows.Add(DR);
            }
            //ListShow(String.Format("{0}", e.symbol));
            //繪製台指期k線
            if (e.symbol == stockCode[0])
            {
                try
                {
                    cal_k(e);
                }
                catch(Exception ex)
                {
                    Debug.WriteLine(ex);
                }                
            }
            //計算力道
            try
            {
                cal_force(e);
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex);
            }            
        }

        public void cal_force(AxYuantaQuoteLib._DYuantaQuoteEvents_OnGetMktAllEvent e)
        {
            int stockIndex = stockCode.IndexOf(e.symbol);
            //初始化stockLastAmount
            if (stockLastAmount[stockIndex] == 0)
            {
                stockLastAmount[stockIndex] = Int32.Parse(e.tolMatchQty);
            }
            //判斷成交量是否改變
            if (Convert.ToString(stockLastAmount[stockIndex]) == e.tolMatchQty)
            {
                return;
            }
            if (stockIndex != -1)
            {
                //第一次紀錄 初始化stockLastPrice
                if (stockLastPrice[stockIndex] == 0)
                {
                    stockLastPrice[stockIndex] = Convert.ToDouble(e.matchPri);
                }
                //計算買賣方向是否改變
                if (Convert.ToDouble(e.matchPri) > stockLastPrice[stockIndex])
                {
                    stockState[stockIndex] = 1;
                }
                else if (Convert.ToDouble(e.matchPri) < stockLastPrice[stockIndex])
                {
                    stockState[stockIndex] = -1;
                }
                //計算該tick成交數量
                int Q = Int32.Parse(e.tolMatchQty) - stockLastAmount[stockIndex];
                //套用至stockTimeForce每秒鐘買賣力道
                try
                {
                    stockTimeForce[stockIndex][Int32.Parse(e.matchTime.Substring(4, 2))] += stockState[stockIndex] * Q;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Time Data Error:" + ex);
                }
                //更新成交量與成交價
                try
                {
                    stockLastPrice[stockIndex] = Convert.ToDouble(e.matchPri);
                    stockLastAmount[stockIndex] = Int32.Parse(e.tolMatchQty);
                    cal_result(DateTime.Now.Second);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine("Data Error:" + ex);
                }
            }
        }

        public void cal_result(int seconds)
        {
            //計算權重
            //從現在秒數前一分鐘算起
            seconds = seconds - 1;
            //重置stockForce
            for (int i = 0; i < stockTimeForce.Count; i++)
            {
                stockForce[i] = 0;
            }
            //計算每一商品的force
            for (int i = 0; i < stockCode.Count; i++)
            {
                //Handle秒數問題
                if (seconds < calInterval)
                {
                    int second_60_change = 60 - calInterval + seconds;
                    for (int j = 59; j >= second_60_change; j--)
                    {
                        stockForce[i] += stockTimeForce[i][j];

                    }
                    for (int j = seconds; j >= 0; j--)
                    {
                        stockForce[i] += stockTimeForce[i][j];
                    }
                }
                else
                {
                    for (int j = seconds; j > seconds - calInterval; j--)
                    {
                        stockForce[i] += stockTimeForce[i][j];
                    }
                }
                DataRow DR = this.dt.Rows.Find(stockCode[i]);
                DR["買賣力道"] = stockForce[stockCode.IndexOf(stockCode[i])];
            }
            //計算權值股權重
            WeightSum = 0;
            for (int i = 1; i < stockCode.Count; i++)
            {
                WeightSum += 0.01 * stockWeight[i] * stockForce[i];
            }
            //計算台指期權重
            txfWeight = 0;
            txfWeight = 0.01 * stockWeight[0] * stockForce[0];
            //寫入檔案
            Debug.WriteLine("WeightSum:" + WeightSum);
            Debug.WriteLine("txfWeight:" + txfWeight);
            File.AppendAllText("D://log.txt", "\nTime:" + DateTime.Now.ToString("hh:mm:ss"));
            File.AppendAllText("D://log.txt", "\nSum:" + WeightSum.ToString("F6"));
            File.AppendAllText("D://log.txt", "\nTXF:" + txfWeight.ToString("F6"));
            File.AppendAllText("D://log.txt", "\nPrice:" + stockLastPrice[0].ToString("F6"));
        }

        public void read50stock()
        {
            using (var fs = File.OpenRead(System.IO.Directory.GetCurrentDirectory()+"\\50stock.csv"))
            using (var reader = new StreamReader(fs))
            {
                int countStock = 0;
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    stockCode.Add(values[1]);
                    stockWeight.Add(Convert.ToDouble(values[3].TrimEnd('%')));
                    stockForce.Add(0);
                    stockState.Add(1);
                    stockLastPrice.Add(0);
                    stockLastAmount.Add(0);
                    //初始化每一商品60秒鐘每一個值
                    stockTimeForce.Add(new List<int>());
                    for (int i = 0; i < 60; i++)
                    {
                        stockTimeForce[countStock].Add(0);                        
                    }
                    countStock++;
                }
            }
        }

        public void load50stock()
        {
            for (int i = 0; i < stockCode.Count; i++)
            {
                axYuantaQuote1.AddMktReg(stockCode[i], "4");
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
                draw_k();
            }
            //判斷是否新的一分鐘
            if (String.Compare(e.matchTime.Substring(2, 2), minutes, true) != 0)
            {
                mclose = lastprice;
                /*TODO:開始畫圖*/
                popen.Add(mopen);
                pclose.Add(mclose);
                phigh.Add(mhigh);
                plow.Add(mlow);
                draw_k();
                //初始化該分鐘的資料
                mhigh = tmp;
                mlow = tmp;
                mopen = tmp;
                mclose = tmp;
                minutes = e.matchTime.Substring(2, 2);
            }
            //比較高低價
            if (tmp > mhigh)
            {
                mhigh = tmp;
            }
            if (tmp < mlow)
            {
                mlow = tmp;
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
            for (int i = shift; i < popen.Count; i++)
            {
                int kxWidth = 5;
                kx = kx - kxWidth * scale;
                double highy = (cs.Height / (max - min)) * (max - phigh[i]);
                double lowy = (cs.Height / (max - min)) * (max - plow[i]);
                double openy = (cs.Height / (max - min)) * (max - popen[i]);
                double closey = (cs.Height / (max - min)) * (max - pclose[i]);
                if (popen[i] > pclose[i])//跌
                {
                    Rectangle rect = new Rectangle(kx, (int)openy, kxWidth * 2, (int)(closey - openy));
                    g.DrawLine(penGreen2, kx + kxWidth, (float)highy, kx + kxWidth, (float)lowy);
                    g.DrawRectangle(penGreen1, rect);
                    g.FillRectangle(greenBrush, rect);
                }
                else//漲
                {
                    Rectangle rect = new Rectangle(kx, (int)closey, kxWidth * 2, (int)(openy - closey));
                    g.DrawLine(penRed2, kx + kxWidth, (float)highy, kx + kxWidth, (float)lowy);
                    g.DrawRectangle(penRed1, rect);
                    g.FillRectangle(redBrush, rect);
                }
                kx = kx - kxWidth * 2 - 2;
            }
        }

        private void button_login_Click(object sender, EventArgs e)
        {

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
               // int RegErrCode = axYuantaQuote1.AddMktReg(textBox_sym.Text.Trim(), "4");
                //textBox_status2.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + RegErrCode.ToString();
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
                //int RegErrCode = axYuantaQuote1.DelMktReg(textBox_sym.Text.Trim());
               // textBox_status2.Text = RegErrCode.ToString();
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
            DataColumn Col14 = new DataColumn("買賣力道", System.Type.GetType("System.String"));
            DataColumn Col15 = new DataColumn("權重", System.Type.GetType("System.String"));
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
            this.dt.Columns.Add(Col14);
            this.dt.Columns.Add(Col15);
            this.dt.PrimaryKey = new DataColumn[] { this.dt.Columns["商品代碼"] };
            bindingSource1.DataSource = this.dt;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            LoginFn();
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

                    if ((this.DataGrid.Columns["OpenPri"].Index == e.ColumnIndex && e.RowIndex >= 0) || (this.DataGrid.Columns["HighPri"].Index == e.ColumnIndex && e.RowIndex >= 0) || (this.DataGrid.Columns["LowPri"].Index == e.ColumnIndex && e.RowIndex >= 0) || (this.DataGrid.Columns["MatchPri"].Index == e.ColumnIndex && e.RowIndex >= 0))
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
            {
                Debug.WriteLine("Error:" + ex);
            }
        }

        private void textBox_id_Click(object sender, EventArgs e)
        {
            textBox_id.SelectAll();
        }

        private void textBox_sym_Click(object sender, EventArgs e)
        {
            //textBox_sym.SelectAll();
        }

        private delegate void InvokeFunction(string msg);
        public void ListShow(string str_log)
        {
            string StrLog = string.Format("{0}  [{1}] ", DateTime.Now.ToString("HH:mm:ss "), str_log);
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

        private void OnTimedEvent(Object source, System.Timers.ElapsedEventArgs e)
        {            
            //Console.WriteLine("The Elapsed event was raised at {0}", e.SignalTime);
            int seconds = DateTime.Now.Second;
            //調整秒差     
            int calLowseconds = seconds - calInterval;
            int newcalseconds = 0;
            if (calLowseconds < 0)
            {
                newcalseconds = calLowseconds + 60;
            }
            //掃描每一檔商品
            for (int n = 0; n < stockCode.Count; n++)
            {
                //判斷秒數相減是否小於0
                if (calLowseconds < 0)
                {
                    for (int i = 0; i < 60; i++)
                    {
                        if (i > seconds && i < newcalseconds)
                        {
                            stockTimeForce[n][i] = 0;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < 60; i++)
                    {
                        if (i > seconds || i < calLowseconds)
                        {
                            stockTimeForce[n][i] = 0;
                        }
                    }
                }
            }
        }

        private void OnTimedEventb(Object source, System.Timers.ElapsedEventArgs e)
        {
            //交易策略撰寫    
            //策略一
            //買賣力道不同向 開始買賣台指期            
            //判斷是否滿足策略一條件-力道反向
            if (WeightSum * txfWeight < 0)
            {
                //判斷是否觸發策略一買賣力道差距條件
                if (Math.Abs(WeightSum - txfWeight) > strategy1_BuySellLength)
                {
                    //判斷台指期交易方向
                    if (WeightSum > txfWeight)
                    {
                        //買台指期
                        tradeDIR = 1;
                    }
                    else if(WeightSum < txfWeight)
                    {
                        //賣台指期
                        tradeDIR = -1;
                    }
                    //判斷是否設定自動交易
                    if (strategy1_autoTrade)
                    {
                        //判斷是否超過投資預算
                        if (initMoney < stockLastPrice[0])
                        {
                            //textBox_status.Text = "你沒有錢 哈哈哈哈";
                        }
                        else
                        {
                            //執行本次交易
                            tradeTXF();
                            strategy1_autoTrade = false;
                        }
                    }
                }
                else
                {
                    //未觸發買賣條件 不買賣
                    tradeDIR = 0;
                }
            }

            //策略二
            //買賣力道同向  計算出力道差距的倍率決定是否購買
            //買賣力道同向
            if (WeightSum * txfWeight > 0)
            {
                //判斷是否觸發設定的倍率
                if ((WeightSum / txfWeight) > strategy2_BuySellmultiply || (txfWeight / WeightSum) > strategy2_BuySellmultiply)
                {
                    if (WeightSum > txfWeight)
                    {
                        //買台指期
                        tradeDIR = 1;
                    }
                    else if(WeightSum < txfWeight)
                    {
                        //賣台指期
                        tradeDIR = -1;
                    }
                    //判斷是否設定自動交易
                    if (strategy2_autoTrade)
                    {
                        //判斷是否超過投資預算
                        if (initMoney < stockLastPrice[0])
                        {
                            //textBox_status.Text = "你沒有錢 哈哈哈哈";
                        }
                        else
                        {
                            //執行本次交易
                            tradeTXF();
                            strategy2_autoTrade = false;
                        }
                    }
                }
                else
                {
                    //未觸發買賣條件 不買賣
                    tradeDIR = 0;
                }
            }
            tradeTXFcover();
            stoploss();
            try
            {
                //stoploss();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show(""+ex);
            }            
            try
            {
                update_parameter();
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void OnTimedEventtrade(Object source, System.Timers.ElapsedEventArgs e)
        {
            strategy1_autoTrade = true;
            strategy2_autoTrade = true;
        }

        public void stoploss()
        {
            //看一下之前賣空的台指期 原本預期跌  結果一直漲
            for (int i = 0; i < TXFpriceN.Count; i++)
            {
                if (stockLastPrice[0] - TXFpriceN[i] > penalty)
                {
                    //忍痛買回來      
                    //計算本交易收益
                    totalAward -= stockLastPrice[0] - TXFpriceN[i] + fee;
                    ListShow("停損-買回一口:賠" + (stockLastPrice[0] - TXFpriceN[i] + fee));
                    try
                    {
                        File.AppendAllText("D://trade.txt", "停損-賣出一口:賠" + (TXFpriceP[i] - stockLastPrice[0] + fee));
                    }
                    catch(Exception ex)
                    {
                        Debug.WriteLine(ex);
                    }
                    initMoney += stockLastPrice[0];
                    TXFamount--;
                    lose++;
                    TXFpriceN.RemoveAt(i);                    
                }
            }
            //看一下之前買多的台指期
            for (int i = 0; i < TXFpriceP.Count; i++)
            {
                if (TXFpriceP[i] - stockLastPrice[0] > penalty)
                {
                    //忍痛賣出                 
                    //計算本交易收益
                    totalAward -= TXFpriceP[i] - stockLastPrice[0] + fee;
                    ListShow("停損-賣出一口:賠" + (TXFpriceP[i] - stockLastPrice[0] + fee));
                    File.AppendAllText("D://trade.txt", "停損-賣出一口:賠" + (TXFpriceP[i] - stockLastPrice[0] + fee));
                    initMoney += stockLastPrice[0];
                    TXFamount--;
                    lose++;
                    TXFpriceP.RemoveAt(i);                    
                }
            }
        }

        public void tradeTXF()
        {
            if (initMoney < stockLastPrice[0])
            {
                //textBox_status.Text = "你沒有錢 哈哈哈哈";
            }
            else
            {
                if (tradeDIR == 1)
                {
                    //買台指期
                    initMoney -= stockLastPrice[0];
                    TXFamount++;
                    TXFpriceP.Add(stockLastPrice[0]);
                    ListShow("多單一口-成功@"+ stockLastPrice[0]);
                    File.AppendAllText("D://trade.txt", "多單一口-成功@" + stockLastPrice[0]);
                }
                else if (tradeDIR == -1)
                {
                    //賣台指期
                    initMoney -= stockLastPrice[0];
                    TXFamount++;
                    TXFpriceN.Add(stockLastPrice[0]);
                    ListShow("空單一口-成功@" + stockLastPrice[0]);
                    File.AppendAllText("D://trade.txt", "空單一口-成功@" + stockLastPrice[0]);
                }
                else if (tradeDIR == 0)
                {
                    textBox_status.Text = "現在不是買賣的時候啦!!!";
                }
            }                
        }

        public void tradeTXFcover()
        {
            for (int i = 0; i < TXFpriceN.Count; i++)
            {
                if (TXFpriceN[i] - stockLastPrice[0] > fee + award)
                {
                    //買回台指期
                    //計算本交易收益
                    totalAward += TXFpriceN[i] - stockLastPrice[0] - fee;
                    ListShow("買回一口-成功賺@" + (TXFpriceN[i] - stockLastPrice[0] - fee)+"賣價@"+ TXFpriceN[i]+"買價@"+ stockLastPrice[0]);
                    File.AppendAllText("D://trade.txt", "買回一口-成功賺@" + (TXFpriceN[i] - stockLastPrice[0] - fee) + "賣價@" + TXFpriceN[i] + "買價@" + stockLastPrice[0]);
                    initMoney += stockLastPrice[0];
                    TXFamount--;
                    win++;
                    TXFpriceN.RemoveAt(i);                    
                }
            }
            for (int i = 0; i < TXFpriceP.Count; i++)
            {
                if (stockLastPrice[0] - TXFpriceP[i] > fee + award)
                {
                    //賣回台指期
                    //計算本交易收益
                    totalAward += stockLastPrice[0] - TXFpriceP[i] - fee;
                    ListShow("賣回一口-成功賺@" + (stockLastPrice[0] - TXFpriceP[i] - fee) + "賣價@" + stockLastPrice[0] + "買價@" + TXFpriceP[i]);
                    File.AppendAllText("D://trade.txt", "賣回一口-成功賺@" + (stockLastPrice[0] - TXFpriceP[i] - fee) + "賣價@" + stockLastPrice[0] + "買價@" + TXFpriceP[i]);
                    initMoney += stockLastPrice[0];
                    TXFamount--;
                    win++;
                    TXFpriceP.RemoveAt(i);                    
                }
            }
            if (tradeDIR == 0)
            {
                //textBox_status.Text = "現在不是出脫的時候啦!!!";
            }
        }

        private void button_sec_Click(object sender, EventArgs e)
        {
            //執行交易
            if (tradeDIR != 0)
            {
                tradeTXFcover();
            }
        }
        
        private void button_login_Click_1(object sender, EventArgs e)
        {
            //登入後讀取50檔權值股
            LoginFn();
            read50stock();
        }
        
        private void btn_Load_Click_1(object sender, EventArgs e)
        {
            try
            {
                //讀取計算權重週期秒數
                calInterval = Int32.Parse(textBox_sec1.Text.Trim());                
            }
            catch(Exception ex)
            {
                MessageBox.Show(DateTime.Now.ToString("HH:mm:ss.fff ") + "計算秒數請輸入0~59");
                Debug.WriteLine("calInterval Input invalid:"+ex);
                return;
            }
            load_parameter();
            //寫入開始時間
            File.AppendAllText("D://log.txt", "\nStart Time:" + DateTime.Now.ToString("h:mm:ss "));
            File.AppendAllText("D://tick.txt", "\nStart Time:" + DateTime.Now.ToString("h:mm:ss "));
            //設定個股買賣力道區間計算計時器
            aTimer30 = new System.Timers.Timer(timeInterval * 1000);
            aTimer30.Elapsed += OnTimedEvent;
            aTimer30.AutoReset = true;
            aTimer30.Enabled = true;
            //登陸50檔權值股資料            
            load50stock();
            //設定整體買賣力道計時器
            bTimer = new System.Timers.Timer(btimeInterval * 1000);
            bTimer.Elapsed += OnTimedEventb;
            bTimer.AutoReset = true;
            bTimer.Enabled = true;
            //設定交易計時器
            tradeTimer = new System.Timers.Timer(tradetimeInterval * 1000);
            tradeTimer.Elapsed += OnTimedEventtrade;
            tradeTimer.AutoReset = true;
            tradeTimer.Enabled = true;
        }
        
        public void update_parameter()
        {
            Form1.CheckForIllegalCrossThreadCalls = false;
            label_force1.Text = Convert.ToString(txfWeight);
            label_force2.Text = Convert.ToString(WeightSum);
            label_quan.Text = Convert.ToString(TXFamount);
            label_money.Text = Convert.ToString(initMoney);
            double ratio = 0;
            if ((win + lose) != 0)
            {
                ratio = (win / (win + lose));
            }
            else
            {
                ratio = 0;
            }
            label_winratio.Text = Convert.ToString(ratio);
            if (tradeDIR == 1)
            {
                label_suggest.Text = Convert.ToString("買");
                label_suggest.ForeColor = Color.FromArgb(255, 0, 0);
            }
            else if(tradeDIR == -1)
            {
                label_suggest.Text = Convert.ToString("賣");
                label_suggest.ForeColor = Color.FromArgb(0, 255, 0);
            }
            else
            {
                label_suggest.Text = Convert.ToString("去睡覺");
                label_suggest.ForeColor = Color.FromArgb(0, 0, 255);
            }
            if (strategy1_autoTrade)
            {
                btn_str1.Text = "停用策略一";
                btn_str1.BackColor = Color.FromArgb(255, 0, 0);
            }
            else
            {
                btn_str1.Text = "啟用策略一";
                btn_str1.BackColor = Color.FromArgb(0, 255, 0);
            }
            if(strategy2_autoTrade)
            {
                btn_str2.Text = "停用策略二";
                btn_str2.BackColor = Color.FromArgb(255, 0, 0);
            }
            else
            {
                btn_str2.Text = "啟用策略二";
                btn_str2.BackColor = Color.FromArgb(0, 255, 0);
            }
        }

        public void load_parameter()
        {
            calInterval = Int32.Parse(textBox_sec1.Text.Trim());
            tradetimeInterval = Int32.Parse(textBox_sec2.Text.Trim());
            initMoney = Double.Parse(textBox_init.Text.Trim());
            strategy1_BuySellLength = Int32.Parse(textBox_str1.Text.Trim());
            strategy2_BuySellmultiply = Int32.Parse(textBox_str2.Text.Trim());
            fee = Int32.Parse(textBox_fee.Text.Trim());
            award = Int32.Parse(textBox_stop1.Text.Trim());
            penalty = Int32.Parse(textBox_stop2.Text.Trim());            
        }

        private void btn_sec_cover_Click(object sender, EventArgs e)
        {

        }

        private void button_sec_cover_Click(object sender, EventArgs e)
        {

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

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

        private void btn_Writefile_Click(object sender, EventArgs e)
        {

        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void btn_Writefile_Click_1(object sender, EventArgs e)
        {
            tradeTXF();
        }

        private void groupBox2_Enter_1(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged_2(object sender, EventArgs e)
        {

        }

        private void axYuantaQuote1_OnMktStatusChange_1(object sender, AxYuantaQuoteLib._DYuantaQuoteEvents_OnMktStatusChangeEvent e)
        {
            textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + e.msg.ToString();
            if (e.msg.ToString().IndexOf("行情連線結束") >= 0)
            {
                //隔幾秒再連線
                textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線結束，隔5秒重新連線";
                MessageBox.Show(DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線結束，隔5秒重新連線");
                timer1.Enabled = true;
            }
            else if (e.msg.ToString().IndexOf("行情連線失敗") >= 0)
            {
                //隔幾秒再連線
                //可能網路不通
                textBox_status.Text = DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線失敗，隔5秒重新連線";
                MessageBox.Show(DateTime.Now.ToString("HH:mm:ss.fff ") + "行情連線失敗，隔5秒重新連線");
                timer1.Enabled = true;
            }
        }

        private void btn_donate_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.twitch.tv/jonathanlu0405");
        }

        private void btn_str1_Click(object sender, EventArgs e)
        {
            if (strategy1_autoTrade)
            {
                strategy1_autoTrade = false;
            }
            else
            {
                strategy1_autoTrade = true;
            }
            update_parameter();
        }

        private void btn_str2_Click(object sender, EventArgs e)
        {
            if (strategy2_autoTrade)
            {
                strategy2_autoTrade = false;
            }
            else
            {
                strategy2_autoTrade = true;
            }
            update_parameter();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void textBox_status2_TextChanged(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }
    }
}
