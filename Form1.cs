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
using System.Windows.Forms.DataVisualization.Charting;

namespace Statistic
{
    public partial class Form1 : Form
    {
        int num;
        int range_num;
        int[] nmax = new int [3];
        int[] nmin = new int [3];
        bool cflg;
        int[] SN = new int[1000000];
        int[] CODE = new int[1000000];
        double[] ID = new double[1000000];
        double[] CS = new double[1000000];
        double[] OD = new double[1000000];

        int[] ID_count = new int[10000];
        int[] CS_count = new int[10000];
        int[] OD_count = new int[10000];

        double[] ID_pos = new double[10000];
        double[] CS_pos = new double[10000];
        double[] OD_pos = new double[10000];

        double[] max = new double[3];
        double[] min = new double[3];
        double[] sum = new double[3];
        double[] AVE = new double[3];
        double[] sum_sqrt = new double[3];
        double[] sqrt_sum_n = new double[3];
        double[] ss = new double[3];
        double[] variance = new double[3];
        double[] STD = new double[3];
        double[] range_scale = new double[3];

        double[] USL = new double[3];
        double[] LSL = new double[3];
        double[] U = new double[3];
        double[] T = new double[3];
        double[] Cpk = new double[3];
        double[] Cp = new double[3];
        double[] Ck = new double[3];

        double[] rmax = new double[3];
        double[] rmin = new double[3];

        double[] max_L = new double[3];
        double[] min_L = new double[3];

        double[] fix = new double[3];


        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            Panel pn = (Panel)sender;

            ControlPaint.DrawBorder(e.Graphics, pn.ClientRectangle,
            Color.White, 1, ButtonBorderStyle.Solid, //左邊
            Color.White, 1, ButtonBorderStyle.Solid, //上邊
            Color.DimGray, 1, ButtonBorderStyle.Solid, //右邊
            Color.DimGray, 1, ButtonBorderStyle.Solid);//底邊
        }
        // 讀文件
        public void ReadFile(string filename)
        {
            if (!File.Exists(@filename))
            {
                MessageBox.Show("檔案不存在，請重新確認!");
            }
            else
            {
                using (StreamReader sr = File.OpenText(@filename))
                {
                    string s = "";
                    string res = "";
                    int pos;

                    num = 0;
                    s = sr.ReadLine();
                    while ((s = sr.ReadLine()) != null)
                    {
                        pos = s.IndexOf(",");
                        res = s.Substring(0, pos);
                        s = s.Substring(pos+1, s.Length - (pos+1));
                        SN[num] = Convert.ToInt32(res);

                        pos = s.IndexOf(",");
                        res = s.Substring(0, pos);
                        s = s.Substring(pos + 1, s.Length - (pos + 1));
                        CODE[num] = Convert.ToInt32(res);

                        pos = s.IndexOf(",");
                        res = s.Substring(0, pos);
                        s = s.Substring(pos + 1, s.Length - (pos + 1));
                        ID[num] = Convert.ToDouble(res);
                        //ID[num] = Convert.ToDouble(res) +0.03; //
                        //MessageBox.Show(res);

                        pos = s.IndexOf(",");
                        res = s.Substring(0, pos);
                        s = s.Substring(pos + 1, s.Length - (pos + 1));
                        CS[num] = Convert.ToDouble(res);
                        //CS[num] = Convert.ToDouble(res) + 0.01; //ku
                        //MessageBox.Show(res);

                        pos = s.IndexOf(",");
                        res = s.Substring(0, pos);
                        s = s.Substring(pos + 1, s.Length - (pos + 1));
                        OD[num] = Convert.ToDouble(res);
                        //MessageBox.Show(res);

                        //if (CODE[num] == 1)
                        //{
                            dataGridView1.Rows.Add(SN[num], ID[num], CS[num], OD[num], CODE[num]);

                            num++;
                        //}
                    }
                    
                    //MessageBox.Show(num.ToString());
                    sr.Close();
                }
            }
        }

        public void CalStatistic()
        {
            int i;
            double[] bmax = new double[3];
            double[] bmin = new double[3];
            // max[0]=max id,max[1]=max cs max[2]=max od 同理
            for (i = 0; i < 3; i++)
            {
                max[i] = -1e10; // 測量的最大值
                min[i] = 1e10;  // 測量最小值
                sum[i] = 0.0;   // 合計 ID
                sum_sqrt[i] = 0.0; // 平方加總
            }


            for (i = 0; i < num; i++)
            {
                if (ID[i] > max[0])
                    max[0] = ID[i];
                if (CS[i] > max[1])
                    max[1] = CS[i];
                if (OD[i] > max[2])
                    max[2] = OD[i];

                if (ID[i] < min[0])
                    min[0] = ID[i];
                if (CS[i] < min[1])
                    min[1] = CS[i];
                if (OD[i] < min[2])
                    min[2] = OD[i];

                sum[0] = sum[0] + ID[i];
                sum[1] = sum[1] + CS[i];
                sum[2] = sum[2] + OD[i];

                sum_sqrt[0] = sum_sqrt[0] + ID[i] * ID[i];
                sum_sqrt[1] = sum_sqrt[1] + CS[i] * CS[i];
                sum_sqrt[2] = sum_sqrt[2] + OD[i] * OD[i];
            }

            AVE[0] = sum[0] / num;  // ID平均值
            AVE[1] = sum[1] / num;  // CS平均值
            AVE[2] = sum[2] / num;  // OD 平均值

            sqrt_sum_n[0] = sum[0] * sum[0] / num;
            sqrt_sum_n[1] = sum[1] * sum[1] / num;
            sqrt_sum_n[2] = sum[2] * sum[2] / num;

            ss[0] = sum_sqrt[0] - sqrt_sum_n[0];
            ss[1] = sum_sqrt[1] - sqrt_sum_n[1];
            ss[2] = sum_sqrt[2] - sqrt_sum_n[2];
            
            variance[0] = ss[0] / num;
            variance[1] = ss[1] / num;
            variance[2] = ss[2] / num;

            STD[0] = Math.Sqrt(variance[0]);
            STD[1] = Math.Sqrt(variance[1]);
            STD[2] = Math.Sqrt(variance[2]);

            T[0] = USL[0] - LSL[0];
            T[1] = USL[1] - LSL[1];
            T[2] = USL[2] - LSL[2];

            U[0] = (USL[0] + LSL[0]) / 2.0;
            U[1] = (USL[1] + LSL[1]) / 2.0;
            U[2] = (USL[2] + LSL[2]) / 2.0;

            Cp[0] = T[0] / (6*STD[0]);
            Cp[1] = T[1] / (6 * STD[1]);
            Cp[2] = T[2] / (6 * STD[2]);

            Ck[0] = (AVE[0] - U[0]) / (T[0] / 2.0);
            Ck[1] = (AVE[1] - U[1]) / (T[1] / 2.0);
            Ck[2] = (AVE[2] - U[2]) / (T[2] / 2.0);

            Cpk[0] = Cp[0] * (1 - Math.Abs(Ck[0]));
            Cpk[1] = Cp[1] * (1 - Math.Abs(Ck[1]));
            Cpk[2] = Cp[2] * (1 - Math.Abs(Ck[2]));

            range_num = Convert.ToInt32( Math.Round(1 + 3.322 * Math.Log10(num), 0, MidpointRounding.AwayFromZero) );
            //range_num = 12; //ku

            bmax[0] = USL[0];
            if (bmax[0] < max[0])
                bmax[0] = max[0];

            bmin[0] = LSL[0];
            if (bmin[0] > min[0])
                bmin[0] = min[0];

            rmax[0] = Math.Round(bmax[0] + (bmax[0] - bmin[0]) * 0.1, 3, MidpointRounding.AwayFromZero);
            rmax[1] = Math.Round(max[1] + (max[1] - min[1]) * 0.05, 3, MidpointRounding.AwayFromZero);
            rmax[2] = Math.Round(max[2] + (max[2] - min[2]) * 0.05, 3, MidpointRounding.AwayFromZero);


            rmin[0] = Math.Round(bmin[0] - (bmax[0] - bmin[0]) * 0.1, 3, MidpointRounding.AwayFromZero);
            rmin[1] = Math.Round(min[1] - (max[1] - min[1]) * 0.05, 3, MidpointRounding.AwayFromZero);
            rmin[2] = Math.Round(min[2] - (max[2] - min[2]) * 0.05, 3, MidpointRounding.AwayFromZero);

            nmax[0] = num / 3;
            nmin[0] = 0;

            nmax[1] = num / 3;
            nmin[1] = 0;

            nmax[2] = num / 3;
            nmin[2] = 0;

            textBox9.Text = rmax[0].ToString();
            textBox12.Text = rmax[1].ToString();
            textBox14.Text = rmax[2].ToString();

            textBox10.Text = rmin[0].ToString();
            textBox11.Text = rmin[1].ToString();
            textBox13.Text = rmin[2].ToString();

            textBox16.Text = nmax[0].ToString();
            textBox15.Text = nmin[0].ToString();

            textBox23.Text = nmax[1].ToString();
            textBox8.Text = nmin[1].ToString();

            textBox27.Text = nmax[2].ToString();
            textBox26.Text = nmin[2].ToString();

            textBox18.Text = max[0].ToString();
            textBox20.Text = max[1].ToString();
            textBox22.Text = max[2].ToString();

            textBox17.Text = min[0].ToString();
            textBox19.Text = min[1].ToString();
            textBox21.Text = min[2].ToString();



            richTextBox1.AppendText("*** 內徑統計資料 ***\n");
            richTextBox1.AppendText("1. 總數量：" + num.ToString() + "\n");
            richTextBox1.AppendText("2. 最大值：" + Math.Round(max[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("3. 最小值：" + Math.Round(min[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("4. 總  和：" + Math.Round(sum[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("5. 平均值：" + Math.Round(AVE[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("6. 標準差：" + Math.Round(STD[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("7. 準度(Ck)：" + Math.Round(Ck[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("8. 精度(Cp)：" + Math.Round(Cp[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("9. 精度(Cpk)：" + Math.Round(Cpk[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("10. 自然公差：" + Math.Round(AVE[0]- 3 * STD[0], 3, MidpointRounding.AwayFromZero).ToString() + " ~ " + Math.Round(AVE[0] + 3 * STD[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");



            richTextBox1.AppendText("\n\n*** 線徑統計資料 ***\n");
            richTextBox1.AppendText("1. 總數量：" + num.ToString() + "\n");
            richTextBox1.AppendText("2. 最大值：" + Math.Round(max[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("3. 最小值：" + Math.Round(min[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("4. 總  和：" + Math.Round(sum[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("5. 平均值：" + Math.Round(AVE[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("6. 標準差：" + Math.Round(STD[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("7. 準度(Ck)：" + Math.Round(Ck[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("8. 精度(Cp)：" + Math.Round(Cp[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("9. 精度(Cpk)：" + Math.Round(Cpk[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("10. 自然公差：" + Math.Round(AVE[1] - 3 * STD[1], 3, MidpointRounding.AwayFromZero).ToString() + " ~ " + Math.Round(AVE[1] + 3 * STD[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");

            richTextBox1.AppendText("\n\n*** 外徑統計資料 ***\n");
            richTextBox1.AppendText("1. 總數量：" + num.ToString() + "\n");
            richTextBox1.AppendText("2. 最大值：" + Math.Round(max[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("3. 最小值：" + Math.Round(min[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("4. 總  和：" + Math.Round(sum[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("5. 平均值：" + Math.Round(AVE[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("6. 標準差：" + Math.Round(STD[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("7. 準度(Ck)：" + Math.Round(Ck[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("8. 精度(Cp)：" + Math.Round(Cp[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("9. 精度(Cpk)：" + Math.Round(Cpk[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox1.AppendText("10. 自然公差：" + Math.Round(AVE[2] - 3 * STD[2], 3, MidpointRounding.AwayFromZero).ToString() + " ~ " + Math.Round(AVE[2] + 3 * STD[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
        }

        public void ReadPara( string path )
        {
            path = path.Substring(0, path.LastIndexOf("\\") + 1);
            path = path + "report.txt";

            if (!File.Exists(@path))
            {
                MessageBox.Show("Report檔案不存在，請重新確認!");
            }
            else
            {
                LSL[0] = 54.15;
                USL[0] = 53.13;
                LSL[1] = 2.7;
                USL[1] = 2.54;
                LSL[2] = 59.55;
                USL[2] = 58.21;


 //                Mean inside diameter:54.15mm
 //Mean inside diameter:53.13mm
 //Mean material cross-section:2.7mm
 //Mean material cross-section:2.54mm
 //Mean outside diameter:59.55mm
 //Mean outside diameter:58.21mm
                //using (StreamReader sr = File.OpenText(@path))
                //{
                //    string s = "";
                //    string res = "";
                //    int pos1, pos2;
                //    double dtmp;
                //    bool flg = true;

                //    s = sr.ReadLine();
                //    while ((s = sr.ReadLine()) != null)
                //    {
                //        if ((s.IndexOf("Mean inside diameter") > 0)&&(flg))
                //        {
                //            s = sr.ReadLine();
                //            pos1 = s.IndexOf(":");
                //            pos2 = s.IndexOf("mm");
                //            res = s.Substring(pos1+1, (pos2-pos1-1) );
                //            dtmp = Convert.ToDouble(res);
                //            LSL[0] = dtmp;  //ID规格上限
                //            //LSL[0] = dtmp + fix[0]; //
                //            textBox3.Text = LSL[0].ToString();

                //            s = sr.ReadLine();
                //            pos1 = s.IndexOf(":");
                //            pos2 = s.IndexOf("mm");
                //            res = s.Substring(pos1 + 1, (pos2 - pos1 - 1));
                //            dtmp = Convert.ToDouble(res);
                //            USL[0] = dtmp; //OD规格下限
                //            //USL[0] = dtmp + fix[0]; //
                //            textBox2.Text = USL[0].ToString();
                //        }

                //        if ((s.IndexOf("Mean material cross-section") > 0) && (flg))
                //        {
                //            s = sr.ReadLine();
                //            pos1 = s.IndexOf(":");
                //            pos2 = s.IndexOf("mm");
                //            res = s.Substring(pos1 + 1, (pos2 - pos1 - 1));
                //            dtmp = Convert.ToDouble(res);
                //            LSL[1] = dtmp;
                //            //LSL[1] = dtmp + fix[1]; //ku
                //            textBox4.Text = LSL[1].ToString();

                //            s = sr.ReadLine();
                //            pos1 = s.IndexOf(":");
                //            pos2 = s.IndexOf("mm");
                //            res = s.Substring(pos1 + 1, (pos2 - pos1 - 1));
                //            dtmp = Convert.ToDouble(res);
                //            USL[1] = dtmp;
                //            //USL[1] = dtmp + fix[1]; //ku
                //            textBox5.Text = USL[1].ToString();
                //        }

                //        if ((s.IndexOf("Mean outside diameter") > 0) && (flg))
                //        {
                //            s = sr.ReadLine();
                //            pos1 = s.IndexOf(":");
                //            pos2 = s.IndexOf("mm");
                //            res = s.Substring(pos1 + 1, (pos2 - pos1 - 1));
                //            dtmp = Convert.ToDouble(res);
                //            LSL[2] = dtmp;
                //            //LSL[2] = dtmp + fix[2];
                //            textBox6.Text = LSL[2].ToString();

                //            s = sr.ReadLine();
                //            pos1 = s.IndexOf(":");
                //            pos2 = s.IndexOf("mm");
                //            res = s.Substring(pos1 + 1, (pos2 - pos1 - 1));
                //            dtmp = Convert.ToDouble(res);
                //            USL[2] = dtmp;
                //            //USL[2] = dtmp + fix[2];
                //            textBox7.Text = USL[2].ToString();

                //            flg = false;
                //        }
                //    }
                //    sr.Close();
                //}

                /*
                fix[0] = 0.0;
                textBox24.Text = fix[0].ToString();
                fix[1] = 0.0;
                textBox25.Text = fix[1].ToString();
                fix[2] = 0.0;
                textBox28.Text = fix[2].ToString();
                 */
                fix[0] = Convert.ToDouble(textBox24.Text);
                fix[1] = Convert.ToDouble(textBox25.Text);
                fix[2] = Convert.ToDouble(textBox28.Text);
            }
        }

        public void PlotResult(object sender1, object sender2, int Nid)
        {
            Chart chart1 = (Chart)sender1;
            PictureBox pb = (PictureBox)sender2;


            int i;
            double hmax = -1e10;
            double cmax = -1e10;
            double cnum = 1000;
            double step;
            double scale;
            double dtmpx;
            double dtmpy;

            chart1.ChartAreas["ChartArea1"].AxisX.Maximum = rmax[Nid];
            chart1.ChartAreas["ChartArea1"].AxisX.Minimum = rmin[Nid];

            chart1.ChartAreas["ChartArea1"].AxisY.Maximum = nmax[Nid];
            chart1.ChartAreas["ChartArea1"].AxisY.Minimum = nmin[Nid];

            chart1.Series["Series1"].Points.Clear();
            for (i = 0; i < range_num; i++)
            {
                if (Nid == 0)
                {
                    if (checkBox1.Checked == true)
                        chart1.Series["Series1"].Points.AddXY(ID_pos[i], ID_count[i]);

                    //richTextBox5.AppendText(ID_count[i].ToString() + "\n");  //
                    //richTextBox5.AppendText(ID_pos[i].ToString() + "\n");  //

                    if (ID_count[i] > hmax)
                        hmax = ID_count[i];
                }
                else if (Nid == 1)
                {
                    if (checkBox16.Checked == true)
                        chart1.Series["Series1"].Points.AddXY(CS_pos[i], CS_count[i]);

                    //richTextBox5.AppendText(CS_count[i].ToString() + "\n");  //ku
                    //richTextBox5.AppendText(CS_pos[i].ToString() + "\n");  //ku

                    if (CS_count[i] > hmax)
                        hmax = CS_count[i];
                }
                else
                {
                    if (checkBox24.Checked == true)
                        chart1.Series["Series1"].Points.AddXY(OD_pos[i], OD_count[i]);

                    if (OD_count[i] > hmax)
                        hmax = OD_count[i];
                }

                
            }

            step = (rmax[Nid] - rmin[Nid]) / cnum;
            for (i = 0; i < cnum; i++)
            {
                dtmpx = rmin[Nid] + step * i;
                dtmpy = (Math.Exp((-1 * (dtmpx - AVE[Nid]) * (dtmpx - AVE[Nid])) / (2 * STD[Nid] * STD[Nid]))) / (STD[Nid] * (Math.Sqrt(2 * Math.PI)));

                if (dtmpy > cmax)
                    cmax = dtmpy;
            }

            chart1.Series["Series2"].Points.Clear();
            scale = hmax / cmax;
            for (i = 0; i < cnum; i++)
            {
                dtmpx = rmin[Nid] + step * i;
                dtmpy = (Math.Exp((-1 * (dtmpx - AVE[Nid]) * (dtmpx - AVE[Nid])) / (2 * STD[Nid] * STD[Nid]))) / (STD[Nid] * (Math.Sqrt(2 * Math.PI)));

                if (Nid == 0)
                {
                    if (checkBox2.Checked == true)
                        chart1.Series["Series2"].Points.AddXY(dtmpx, dtmpy * scale);
                }
                else if (Nid == 1)
                {
                    if (checkBox15.Checked == true)
                        chart1.Series["Series2"].Points.AddXY(dtmpx, dtmpy * scale);
                }
                else 
                {
                    if (checkBox23.Checked == true)
                        chart1.Series["Series2"].Points.AddXY(dtmpx, dtmpy * scale);
                }
            }

            chart1.Series["Series3"].Points.Clear();
            if (Nid == 0)
            {
                if (checkBox3.Checked == true)
                {
                    chart1.Series["Series3"].Points.AddXY(USL[Nid], nmax[Nid]);
                    chart1.Series["Series3"].Points.AddXY(USL[Nid], nmin[Nid]);
                }
            }
            else if (Nid == 1)
            {
                if (checkBox14.Checked == true)
                {
                    chart1.Series["Series3"].Points.AddXY(USL[Nid], nmax[Nid]);
                    chart1.Series["Series3"].Points.AddXY(USL[Nid], nmin[Nid]);
                }
            }
            else
            {
                if (checkBox22.Checked == true)
                {
                    chart1.Series["Series3"].Points.AddXY(USL[Nid], nmax[Nid]);
                    chart1.Series["Series3"].Points.AddXY(USL[Nid], nmin[Nid]);
                }
            }

            chart1.Series["Series4"].Points.Clear();
            if (Nid == 0)
            {
                if (checkBox8.Checked == true)
                {
                    chart1.Series["Series4"].Points.AddXY(LSL[Nid], nmax[Nid]);
                    chart1.Series["Series4"].Points.AddXY(LSL[Nid], nmin[Nid]);
                }
            }
            else if (Nid == 1)
            {
                if (checkBox10.Checked == true)
                {
                    chart1.Series["Series4"].Points.AddXY(LSL[Nid], nmax[Nid]);
                    chart1.Series["Series4"].Points.AddXY(LSL[Nid], nmin[Nid]);
                }
            }
            else
            {
                if (checkBox18.Checked == true)
                {
                    chart1.Series["Series4"].Points.AddXY(LSL[Nid], nmax[Nid]);
                    chart1.Series["Series4"].Points.AddXY(LSL[Nid], nmin[Nid]);
                }
            }

            chart1.Series["Series5"].Points.Clear();
            if (Nid == 0)
            {
                if (checkBox6.Checked == true)
                {
                    chart1.Series["Series5"].Points.AddXY((USL[Nid] + LSL[Nid]) / 2.0, nmax[Nid]);
                    chart1.Series["Series5"].Points.AddXY((USL[Nid] + LSL[Nid]) / 2.0, nmin[Nid]);
                }
            }
            else if (Nid == 1)
            {
                if (checkBox12.Checked == true)
                {
                    chart1.Series["Series5"].Points.AddXY((USL[Nid] + LSL[Nid]) / 2.0, nmax[Nid]);
                    chart1.Series["Series5"].Points.AddXY((USL[Nid] + LSL[Nid]) / 2.0, nmin[Nid]);
                }
            }
            else
            {
                if (checkBox20.Checked == true)
                {
                    chart1.Series["Series5"].Points.AddXY((USL[Nid] + LSL[Nid]) / 2.0, nmax[Nid]);
                    chart1.Series["Series5"].Points.AddXY((USL[Nid] + LSL[Nid]) / 2.0, nmin[Nid]);
                }
            }

            chart1.Series["Series7"].Points.Clear();    //自然公差上限
            if (Nid == 0)
            {
                if (checkBox5.Checked == true)
                {
                    chart1.Series["Series7"].Points.AddXY(AVE[Nid] + 3 * STD[Nid], nmax[Nid]);
                    chart1.Series["Series7"].Points.AddXY(AVE[Nid] + 3 * STD[Nid], nmin[Nid]);
                }
            }
            else if (Nid == 1)
            {
                if (checkBox13.Checked == true)
                {
                    chart1.Series["Series7"].Points.AddXY(AVE[Nid] + 3 * STD[Nid], nmax[Nid]);
                    chart1.Series["Series7"].Points.AddXY(AVE[Nid] + 3 * STD[Nid], nmin[Nid]);
                }
            }
            else
            {
                if (checkBox21.Checked == true)
                {
                    chart1.Series["Series7"].Points.AddXY(AVE[Nid] + 3 * STD[Nid], nmax[Nid]);
                    chart1.Series["Series7"].Points.AddXY(AVE[Nid] + 3 * STD[Nid], nmin[Nid]);
                }
            }

            chart1.Series["Series6"].Points.Clear();
            if (Nid == 0)
            {
                if (checkBox4.Checked == true)
                {
                    chart1.Series["Series6"].Points.AddXY(AVE[Nid], nmax[Nid]);
                    chart1.Series["Series6"].Points.AddXY(AVE[Nid], nmin[Nid]);
                }
            }
            else if (Nid == 1)
            {
                if (checkBox11.Checked == true)
                {
                    chart1.Series["Series6"].Points.AddXY(AVE[Nid], nmax[Nid]);
                    chart1.Series["Series6"].Points.AddXY(AVE[Nid], nmin[Nid]);
                }
            }
            else
            {
                if (checkBox19.Checked == true)
                {
                    chart1.Series["Series6"].Points.AddXY(AVE[Nid], nmax[Nid]);
                    chart1.Series["Series6"].Points.AddXY(AVE[Nid], nmin[Nid]);
                }
            }

            chart1.Series["Series8"].Points.Clear();  //自然公差下限
            if (Nid == 0)
            {
                if (checkBox7.Checked == true)
                {
                    chart1.Series["Series8"].Points.AddXY(AVE[Nid] - 3 * STD[Nid], nmax[Nid]);
                    chart1.Series["Series8"].Points.AddXY(AVE[Nid] - 3 * STD[Nid], nmin[Nid]);
                }
            }
            else if (Nid == 1)
            {
                if (checkBox9.Checked == true)
                {
                    chart1.Series["Series8"].Points.AddXY(AVE[Nid] - 3 * STD[Nid], nmax[Nid]);
                    chart1.Series["Series8"].Points.AddXY(AVE[Nid] - 3 * STD[Nid], nmin[Nid]);
                }
            }
            else
            {
                if (checkBox17.Checked == true)
                {
                    chart1.Series["Series8"].Points.AddXY(AVE[Nid] - 3 * STD[Nid], nmax[Nid]);
                    chart1.Series["Series8"].Points.AddXY(AVE[Nid] - 3 * STD[Nid], nmin[Nid]);
                }
            }

            Bitmap img = new Bitmap(100, 400);
            Graphics bmp = Graphics.FromImage(img);

            Brush b = new SolidBrush(Color.Gray);
            //bmp.FillRectangle(b, 0, 0, 100, 400);

            b = new SolidBrush(Color.Red);
            bmp.FillRectangle(b, 50, 30, 45, 80);
            Pen p = new Pen(Color.Black, 2);
            bmp.DrawRectangle(p, 50, 30, 45, 80);

            b = new SolidBrush(Color.Yellow);
            bmp.FillRectangle(b, 50, 110, 45, 80);
            p = new Pen(Color.Black, 2);
            bmp.DrawRectangle(p, 50, 110, 45, 80);

            b = new SolidBrush(Color.Green);
            bmp.FillRectangle(b, 50, 190, 45, 80);
            p = new Pen(Color.Black, 2);
            bmp.DrawRectangle(p, 50, 190, 45, 80);

            b = new SolidBrush(Color.CornflowerBlue);
            bmp.FillRectangle(b, 50, 270, 45, 80);
            p = new Pen(Color.Black, 2);
            bmp.DrawRectangle(p, 50, 270, 45, 80);


            int xpos = 25;
            int ypos;
            string msg;

            if (Cpk[Nid] < 1.0)
            {
                ypos = 270 + 80 / 2;
                msg = "劣";
            }
            else if (Cpk[Nid] < 1.33)
            {
                ypos = 190 + 80 / 2;
                msg = "差";
            }
            else if (Cpk[Nid] < 2.0)
            {
                ypos = 110 + 80 / 2;
                msg = "良";
            }
            else
            {
                ypos = 30 + 80 / 2;
                msg = "優";
            }

            Point point1 = new Point(xpos - 20, ypos - 16);
            Point point2 = new Point(xpos + 25, ypos - 16);
            Point point3 = new Point(xpos + 50, ypos);
            Point point4 = new Point(xpos + 25, ypos + 16);
            Point point5 = new Point(xpos - 20, ypos + 16);
            Point[] curvePoints = { point1, point2, point3, point4, point5 };



            b = new SolidBrush(Color.Black);
            //FillMode newFillMode = FillMode.Winding;
            bmp.FillPolygon(b, curvePoints);

            b = new SolidBrush(Color.White);
            Font myFont = new Font("宋体", 18, FontStyle.Bold);
            //Brush bush = new SolidBrush(Color.Red);//填充的颜色
            bmp.DrawString(msg, myFont, b, xpos - 15, ypos - 12);

            //bmp.FillPolygon();

            pb.Image = img;
            pb.SizeMode = PictureBoxSizeMode.StretchImage; 
        }

        public void Calculation()
        {
            int i;
            int j;
            bool bflg;
            double[] bmax = new double[3];
            double[] bmin = new double[3];


            //richTextBox5.AppendText("ID: max=" + max_L[0] + ",min=" + min_L[0] + "\n");
            //richTextBox5.AppendText("CS: max=" + max_L[1] + ",min=" + min_L[1] + "\n");
            num = 0;
            for (i = 0; i < dataGridView1.RowCount; i++)
            {
                //if ((Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) >= min_L[0]) && (Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) <= max_L[0]))
                //{
                    //if ((Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) >= min_L[1]) && (Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) <= max_L[1]))
                    //{
                        //if ((Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) >= min_L[2]) && (Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) <= max_L[2]))
                        //{
                            SN[num] = Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value);
                            ID[num] = Convert.ToDouble(dataGridView1.Rows[i].Cells[1].Value) + fix[0];
                            CS[num] = Convert.ToDouble(dataGridView1.Rows[i].Cells[2].Value) + fix[1];
                            OD[num] = Convert.ToDouble(dataGridView1.Rows[i].Cells[3].Value) + fix[2];

                            
                            if ((ID[num] <= max_L[0]) && (ID[num] >= min_L[0]))
                            {
                                if ((CS[num] <= max_L[1]) && (CS[num] >= min_L[1]))
                                {
                                    num++;
                                }
                            }
                        //}
                    //}
                //}
            }


            for (i = 0; i < 3; i++)
            {
                max[i] = -1e10;
                min[i] = 1e10;
                sum[i] = 0.0;
                sum_sqrt[i] = 0.0;
            }


            for (i = 0; i < num; i++)
            {
                if (ID[i] > max[0])
                    max[0] = ID[i];
                if (CS[i] > max[1])
                    max[1] = CS[i];
                if (OD[i] > max[2])
                    max[2] = OD[i];

                if (ID[i] < min[0])
                    min[0] = ID[i];
                if (CS[i] < min[1])
                    min[1] = CS[i];
                if (OD[i] < min[2])
                    min[2] = OD[i];

                sum[0] = sum[0] + ID[i];
                sum[1] = sum[1] + CS[i];
                sum[2] = sum[2] + OD[i];

                sum_sqrt[0] = sum_sqrt[0] + ID[i] * ID[i];
                sum_sqrt[1] = sum_sqrt[1] + CS[i] * CS[i];
                sum_sqrt[2] = sum_sqrt[2] + OD[i] * OD[i];
            }

            AVE[0] = sum[0] / num;
            AVE[1] = sum[1] / num;
            AVE[2] = sum[2] / num;

            sqrt_sum_n[0] = sum[0] * sum[0] / num;
            sqrt_sum_n[1] = sum[1] * sum[1] / num;
            sqrt_sum_n[2] = sum[2] * sum[2] / num;

            ss[0] = sum_sqrt[0] - sqrt_sum_n[0];
            ss[1] = sum_sqrt[1] - sqrt_sum_n[1];
            ss[2] = sum_sqrt[2] - sqrt_sum_n[2];

            variance[0] = ss[0] / num;
            variance[1] = ss[1] / num;
            variance[2] = ss[2] / num;

            STD[0] = Math.Sqrt(variance[0]);
            STD[1] = Math.Sqrt(variance[1]);
            STD[2] = Math.Sqrt(variance[2]);

            T[0] = USL[0] - LSL[0];
            T[1] = USL[1] - LSL[1];
            T[2] = USL[2] - LSL[2];

            U[0] = (USL[0] + LSL[0]) / 2.0;
            U[1] = (USL[1] + LSL[1]) / 2.0;
            U[2] = (USL[2] + LSL[2]) / 2.0;

            Cp[0] = T[0] / (6 * STD[0]);
            Cp[1] = T[1] / (6 * STD[1]);
            Cp[2] = T[2] / (6 * STD[2]);

            Ck[0] = (AVE[0] - U[0]) / (T[0] / 2.0);
            Ck[1] = (AVE[1] - U[1]) / (T[1] / 2.0);
            Ck[2] = (AVE[2] - U[2]) / (T[2] / 2.0);

            Cpk[0] = Cp[0] * (1 - Math.Abs(Ck[0]));
            Cpk[1] = Cp[1] * (1 - Math.Abs(Ck[1]));
            Cpk[2] = Cp[2] * (1 - Math.Abs(Ck[2]));

            

            range_num = Convert.ToInt32(Math.Round(1 + 3.322 * Math.Log10(num), 0, MidpointRounding.AwayFromZero));
            //range_num = 12; //ku

            range_scale[0] = (max[0] - min[0]) / range_num;
            range_scale[1] = (max[1] - min[1]) / range_num;
            range_scale[2] = (max[2] - min[2]) / range_num;
            if (cflg)
            {
                bmax[0] = USL[0];
                if (bmax[0] < max[0])
                    bmax[0] = max[0];

                bmin[0] = LSL[0];
                if (bmin[0] > min[0])
                    bmin[0] = min[0];

                bmax[1] = USL[1];
                if (bmax[1] < max[1])
                    bmax[1] = max[1];

                bmin[1] = LSL[1];
                if (bmin[1] > min[1])
                    bmin[1] = min[1];

                bmax[2] = USL[2];
                if (bmax[2] < max[2])
                    bmax[2] = max[2];

                bmin[2] = LSL[2];
                if (bmin[2] > min[2])
                    bmin[2] = min[2];

                rmax[0] = Math.Round(bmax[0] + (bmax[0] - bmin[0]) * 0.1, 3, MidpointRounding.AwayFromZero);
                rmax[1] = Math.Round(bmax[1] + (bmax[1] - bmin[1]) * 0.05, 3, MidpointRounding.AwayFromZero);
                rmax[2] = Math.Round(bmax[2] + (bmax[2] - bmin[2]) * 0.05, 3, MidpointRounding.AwayFromZero);


                rmin[0] = Math.Round(bmin[0] - (bmax[0] - bmin[0]) * 0.1, 3, MidpointRounding.AwayFromZero);
                rmin[1] = Math.Round(bmin[1] - (bmax[1] - bmin[1]) * 0.05, 3, MidpointRounding.AwayFromZero);
                rmin[2] = Math.Round(bmin[2] - (bmax[2] - bmin[2]) * 0.05, 3, MidpointRounding.AwayFromZero);

                nmax[0] = num / 3;
                nmin[0] = 0;

                nmax[1] = num / 3;
                nmin[1] = 0;

                nmax[2] = num / 3;
                nmin[2] = 0;

                textBox9.Text = rmax[0].ToString();
                textBox12.Text = rmax[1].ToString();
                textBox14.Text = rmax[2].ToString();

                textBox10.Text = rmin[0].ToString();
                textBox11.Text = rmin[1].ToString();
                textBox13.Text = rmin[2].ToString();

                textBox16.Text = nmax[0].ToString();
                textBox15.Text = nmin[0].ToString();

                textBox23.Text = nmax[1].ToString();
                textBox8.Text = nmin[1].ToString();

                textBox27.Text = nmax[2].ToString();
                textBox26.Text = nmin[2].ToString();

                cflg = false;
            }


            for (i = 0; i < range_num; i++)
            {
                ID_count[i] = 0;
                CS_count[i] = 0;
                OD_count[i] = 0;

                ID_pos[i] = Math.Round(min[0] + (i + 0.5) * range_scale[0], 3, MidpointRounding.AwayFromZero);
                CS_pos[i] = Math.Round(min[1] + (i+0.5)*range_scale[1], 3, MidpointRounding.AwayFromZero);
                OD_pos[i] = Math.Round(min[2] + (i + 0.5) * range_scale[2], 3, MidpointRounding.AwayFromZero);
            }

            for (i = 0; i < num; i++)
            {
                j = 0;
                bflg = true;
                do
                {
                    if (ID[i] <=  (min[0] + range_scale[0]*(j+1)))
                    {
                        ID_count[j] = ID_count[j] + 1;
                        bflg = false;
                    }
                    j++;
                }while(bflg);
            }

            for (i = 0; i < num; i++)
            {
                j = 0;
                bflg = true;
                do
                {
                    if (CS[i] <= (min[1] + range_scale[1] * (j + 1)))
                    {
                        CS_count[j] = CS_count[j] + 1;
                        bflg = false;
                    }
                    j++;
                } while (bflg);
            }

            for (i = 0; i < num; i++)
            {
                j = 0;
                bflg = true;
                do
                {
                    if (OD[i] <= (min[2] + range_scale[2] * (j + 1)))
                    {
                        OD_count[j] = OD_count[j] + 1;
                        bflg = false;
                    }
                    j++;
                } while (bflg);
            }

            richTextBox2.Clear();
            richTextBox2.AppendText("\t\t*** 內徑統計資料 ***\n");
            richTextBox2.AppendText("1. 總數量：" + num.ToString() + "\t\t\t" + "7. 準度(Ck)：" + Math.Round(Ck[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox2.AppendText("2. 最大值：" + Math.Round(max[0], 3, MidpointRounding.AwayFromZero).ToString() + "\t\t" + "8. 精度(Cp)：" + Math.Round(Cp[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox2.AppendText("3. 最小值：" + Math.Round(min[0], 3, MidpointRounding.AwayFromZero).ToString() + "\t\t" + "9. 精度(Cpk)：" + Math.Round(Cpk[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox2.AppendText("4. 總  和：" + Math.Round(sum[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox2.AppendText("5. 平均值：" + Math.Round(AVE[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox2.AppendText("6. 標準差：" + Math.Round(STD[0], 3, MidpointRounding.AwayFromZero).ToString() + "\n");

            richTextBox3.Clear();
            richTextBox3.AppendText("\t\t*** 線徑統計資料 ***\n");
            richTextBox3.AppendText("1. 總數量：" + num.ToString() + "\t\t\t" + "7. 準度(Ck)：" + Math.Round(Ck[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox3.AppendText("2. 最大值：" + Math.Round(max[1], 3, MidpointRounding.AwayFromZero).ToString() + "\t\t" + "8. 精度(Cp)：" + Math.Round(Cp[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox3.AppendText("3. 最小值：" + Math.Round(min[1], 3, MidpointRounding.AwayFromZero).ToString() + "\t\t" + "9. 精度(Cpk)：" + Math.Round(Cpk[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox3.AppendText("4. 總  和：" + Math.Round(sum[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox3.AppendText("5. 平均值：" + Math.Round(AVE[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox3.AppendText("6. 標準差：" + Math.Round(STD[1], 3, MidpointRounding.AwayFromZero).ToString() + "\n");

            richTextBox4.Clear();
            richTextBox4.AppendText("\t\t*** 外徑統計資料 ***\n");
            richTextBox4.AppendText("1. 總數量：" + num.ToString() + "\t\t\t" + "7. 準度(Ck)：" + Math.Round(Ck[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox4.AppendText("2. 最大值：" + Math.Round(max[2], 3, MidpointRounding.AwayFromZero).ToString() + "\t\t" + "8. 精度(Cp)：" + Math.Round(Cp[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox4.AppendText("3. 最小值：" + Math.Round(min[2], 3, MidpointRounding.AwayFromZero).ToString() + "\t\t" + "9. 精度(Cpk)：" + Math.Round(Cpk[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox4.AppendText("4. 總  和：" + Math.Round(sum[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox4.AppendText("5. 平均值：" + Math.Round(AVE[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");
            richTextBox4.AppendText("6. 標準差：" + Math.Round(STD[2], 3, MidpointRounding.AwayFromZero).ToString() + "\n");

            PlotResult(chart1, pictureBox1, 0);
            PlotResult(chart2, pictureBox2, 1);
            PlotResult(chart3, pictureBox3, 2);
        }

        public void PlotChart()
        {
            double dtmp;

            nmax[0] = Convert.ToInt32(textBox16.Text);
            nmin[0] = Convert.ToInt32(textBox15.Text);

            nmax[1] = Convert.ToInt32(textBox23.Text);
            nmin[1] = Convert.ToInt32(textBox8.Text);

            nmax[2] = Convert.ToInt32(textBox27.Text);
            nmin[2] = Convert.ToInt32(textBox26.Text);

            rmax[0] = Convert.ToDouble(textBox9.Text);
            rmax[1] = Convert.ToDouble(textBox12.Text);
            rmax[2] = Convert.ToDouble(textBox14.Text);

            rmin[0] = Convert.ToDouble(textBox10.Text);
            rmin[1] = Convert.ToDouble(textBox11.Text);
            rmin[2] = Convert.ToDouble(textBox13.Text);

            USL[0] = Convert.ToDouble(textBox2.Text);
            LSL[0] = Convert.ToDouble(textBox3.Text);
            USL[1] = Convert.ToDouble(textBox5.Text);
            LSL[1] = Convert.ToDouble(textBox4.Text);
            USL[2] = Convert.ToDouble(textBox7.Text);
            LSL[2] = Convert.ToDouble(textBox6.Text);

            dtmp = Convert.ToDouble(textBox18.Text);
            if (max_L[0] != dtmp)
            {
                //MessageBox.Show(dtmp.ToString());
                max_L[0] = dtmp;
                cflg = true;
            }

            dtmp = Convert.ToDouble(textBox20.Text);
            if (max_L[1] != dtmp)
            {
                max_L[1] = dtmp;
                cflg = true;
            }

            dtmp = Convert.ToDouble(textBox22.Text);
            if (max_L[2] != dtmp)
            {
                max_L[2] = dtmp;
                cflg = true;
            }

            dtmp = Convert.ToDouble(textBox17.Text);
            if (min_L[0] != dtmp)
            {
                min_L[0] = dtmp;
                cflg = true;
            }

            dtmp = Convert.ToDouble(textBox19.Text);
            if (min_L[1] != dtmp)
            {
                min_L[1] = dtmp;
                cflg = true;
            }

            dtmp = Convert.ToDouble(textBox21.Text);
            if (min_L[2] != dtmp)
            {
                min_L[2] = dtmp;
                cflg = true;
            }

            fix[0] = Convert.ToDouble(textBox24.Text);
            fix[1] = Convert.ToDouble(textBox25.Text);
            fix[2] = Convert.ToDouble(textBox28.Text);
             

            Calculation();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Application.StartupPath;
            openFileDialog1.Filter = "原始資料 (*.csv)|*.csv|所有檔案 (*.*)|*.*";
            

            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                //System.IO.StreamReader sr = new
                   //System.IO.StreamReader(openFileDialog1.FileName);
                textBox1.Text = openFileDialog1.FileName.ToString();

                ReadFile( openFileDialog1.FileName.ToString() );
                ReadPara( openFileDialog1.FileName.ToString() );
                CalStatistic();

                cflg = false;
            }

        }

        private void panel5_Paint(object sender, PaintEventArgs e)
        {
            Panel pn = (Panel)sender;

            ControlPaint.DrawBorder(e.Graphics, pn.ClientRectangle,
            Color.DimGray, 1, ButtonBorderStyle.Solid, //左邊
            Color.DimGray, 1, ButtonBorderStyle.Solid, //上邊
            Color.White, 1, ButtonBorderStyle.Solid, //右邊
            Color.White, 1, ButtonBorderStyle.Solid);//底邊
        }

        private void button2_Click(object sender, EventArgs e)
        {
            PlotChart();
        }

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            cflg = true;
            button2.PerformClick();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            button2.PerformClick();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.chart1.SaveImage("ID.jpg", ChartImageFormat.Jpeg);
        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }
    }
}
