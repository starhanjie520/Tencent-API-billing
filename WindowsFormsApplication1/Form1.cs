using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TencentCloud.Common;
using TencentCloud.Common.Profile;
using TencentCloud.Billing.V20180709;
using TencentCloud.Billing.V20180709.Models;
using ClosedXML;
using ClosedXML.Excel;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {

        private Thread main_Thread = null;
        public String path = Application.StartupPath + "\\" ;
        public String[] tt = null;
        public int allCompanyIndex = 1;

        public String Company = "";
        public String UIN = "";
        public String SecretIdfromeFile = "";
        public String SecretKeyfromeFile = "";
        public String CVMRegion = "ap-seoul";  //ap-beijing
        public String BeginTime = "";
        public String EndTime = "";
        public String fileName = "";
        public int exelIndex = 2;
        public String months = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            StreamReader sr = new StreamReader("업체.txt", Encoding.UTF8);
            String rrr = sr.ReadToEnd();
            sr.Close();
            tt = Regex.Split(rrr, "\r\n");
            for (int i = 0; i < tt.Count(); i++)
            {
                String CompanyStr = Regex.Split(tt[i], "	")[0];
                checkedListBox1.Items.Add(CompanyStr);
            }


            if (checkBox1.Checked)
            {
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                {
                    checkedListBox1.SetItemChecked(j, true);
                }
            }
            else
            {
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                {
                    checkedListBox1.SetItemChecked(j, false);
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = true;
            main_Thread = new Thread(myStaticThreadMethod);
            main_Thread.Start();

        }

        public void myStaticThreadMethod()
        {
            Test();
            MessageBox.Show("CAPCLOUD 오늘 하루도 수고 하셨습니다. ^_^ ");
            System.Windows.Forms.Application.Exit();
        }


        public void Test()
        {
            int createFileIndex = 1;
            string[] allFiles = Directory.GetFiles(path, "*.xlsx");
            foreach (string strFile in allFiles)
            {
                File.Delete(strFile);
            } 
            ListBoxItemAdd(this, this.listBox1, "프로그램 시작 합니다. ^_^");
            foreach (int indexChecked in checkedListBox1.CheckedIndices)
            {
                String lineStr = tt[indexChecked].ToString();
                Console.WriteLine(lineStr);
                try
                {
                    string[] sArray = lineStr.Split('	');
                    if (sArray.Count() < 2)                                      //작업완료 판단
                    {
                        break;
                    }

                    Company = sArray[0];
                    UIN = sArray[1].Replace(" ","");
                    SecretIdfromeFile = sArray[2].Replace(" ", "");
                    SecretKeyfromeFile = sArray[3].Replace(" ", "");
                    SecretKeyfromeFile = sArray[3].Replace(" ", "");
                    CVMRegion  = sArray[4].Replace(" ", "");

                    BeginTime = dateTimePicker1.Value.ToString("yyyy-MM-01 00:00:00");
                    EndTime = dateTimePicker1.Value.ToString("yyyy-MM-01 00:00:01");
                    months = dateTimePicker1.Value.ToString("yyyy년MM월");
                    Console.WriteLine(Company);
                    Console.WriteLine(UIN);
                    Console.WriteLine(SecretIdfromeFile);
                    Console.WriteLine(SecretKeyfromeFile);
                    Console.WriteLine(BeginTime);
                    Console.WriteLine(EndTime);               
                    Thread.Sleep(1000);
                    ListBoxItemAdd(this, this.listBox1, "======<< 업체명 : " + Company + " >>=====");

                    fileName = createFileIndex + "-" + Company + "-" + months + ".xlsx";   //파일생성
                    WriteExel(fileName, 1, 1, "TotalCost");
                    WriteExel("모든업체-" + months + ".xlsx", allCompanyIndex, 1, Company);

                    Credential cred = new Credential
                    {
                        SecretId = SecretIdfromeFile,
                        SecretKey = SecretKeyfromeFile
                    };

                    ClientProfile clientProfile = new ClientProfile();
                    HttpProfile httpProfile = new HttpProfile();
                    httpProfile.Endpoint = ("billing.tencentcloudapi.com");
                    clientProfile.HttpProfile = httpProfile;
                    BillingClient client = new BillingClient(cred, CVMRegion, clientProfile);
                    DescribeBillSummaryByProductRequest req = new DescribeBillSummaryByProductRequest();
                    string strParams = "{\"PayerUin\":\"" + UIN + "\",\"BeginTime\":\"" + BeginTime + "\",\"EndTime\":\"" + EndTime + "\"}";
                    req = DescribeBillSummaryByProductRequest.FromJsonString<DescribeBillSummaryByProductRequest>(strParams);
                    DescribeBillSummaryByProductResponse resp = client.DescribeBillSummaryByProduct(req).ConfigureAwait(false).GetAwaiter().GetResult();
                    //  Console.WriteLine(AbstractModel.ToJsonString(resp));
                    String ttt = AbstractModel.ToJsonString(resp);

                    String total = GetSummaryTotal(ttt);                 
                    WriteExel(fileName, 1, 2, total);                    
                    WriteExel("모든업체-" + months + ".xlsx", allCompanyIndex, 2, total);
                    allCompanyIndex++;

                    ListBoxItemAdd(this, this.listBox1, "Total : =》  " + total);

                    Dictionary<string, string> Overview = GetSummaryOverview(ttt);
                    exelIndex = 2;
                    foreach (var item in Overview)
                    {
                        Console.WriteLine(item.Key + item.Value);
                        WriteExel(fileName, exelIndex, 1, item.Key);
                        WriteExel(fileName, exelIndex, 2, item.Value);
                        exelIndex++;
                        ListBoxItemAdd(this, this.listBox1, item.Key + "  --->    " +item.Value);
                    }
                    createFileIndex++;
                    Thread.Sleep(100);
                }
                catch (Exception e)
                {
                    //WriteExel(fileName, exelIndex, 1, Company);
                    WriteExel("모든업체-" + months + ".xlsx", allCompanyIndex, 2, "에러");
                    createFileIndex++;
                    ListBoxItemAdd(this, this.listBox1, "<< 업체명 : " + Company + " >> 요금 실패~~");
                }
                ListBoxItemAdd(this, this.listBox1, " ");
                ListBoxItemAdd(this, this.listBox1, " ");
            
            }


        }

        public static string GetSummaryTotal(string jsonText)
        {
            JObject jsonObj = JObject.Parse(jsonText);
            String a = jsonObj["SummaryTotal"]["RealTotalCost"].ToString();
            return a;
        }
        public static Dictionary<string, string> GetSummaryOverview(string jsonText)
        {
            Dictionary<string, string> dic = new Dictionary<string, string>();
            JObject jsonObj = JObject.Parse(jsonText);
            var a = jsonObj.SelectToken("SummaryOverview");
            foreach (var item in a)
            {
                String BusinessCodeName = item.SelectToken("BusinessCodeName").ToString();
                String RealTotalCost = item.SelectToken("RealTotalCost").ToString();
                Console.WriteLine(BusinessCodeName + "-" + RealTotalCost);
                dic.Add(BusinessCodeName, RealTotalCost);
            }
            return dic;
        }
        public void WriteExel(String fileName, int x, int y, String strValue)
        {
            if (!File.Exists(fileName))
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add(fileName.Replace(".xlsx", ""));
                IXLWorksheet sheet = wb.Worksheet(1);
                sheet.Cell(x, y).Value = strValue;
                wb.SaveAs(fileName);
            }
            else
            {
                var wb = new XLWorkbook(fileName);
                IXLWorksheet sheet = wb.Worksheet(1);
                sheet.Cell(x, y).Value = strValue;
                wb.SaveAs(fileName);
            }
        }
      
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            if (this.main_Thread != null)
            {
                this.main_Thread.Abort();
                this.main_Thread = null;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //button1.Enabled = true;
            //button2.Enabled = false;
            if (this.main_Thread != null)
            {
                this.main_Thread.Abort();
                this.main_Thread = null;
            }
        }


        

      
        public static void ListBoxItemAdd(Form frm, ListBox lstbox, string lstitem)
        {
            frm.Invoke(new MethodInvoker(delegate
            {
                if (lstbox.Items.Count + 1 > 200)
                {
                    lstbox.Items.RemoveAt(0);
                }
                lstbox.Items.Add(DateTime.Now.ToString("HH:mm:ss") + " | " + lstitem);
                lstbox.SelectedIndex = lstbox.Items.Count - 1;
                lstbox.Refresh();
                //StreamWriter sw = new StreamWriter(Application.StartupPath + "\\Log.txt", true, Encoding.Default);
                //sw.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " | " + lstitem);
                //sw.Close();
            }));
        }
      
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                {
                    checkedListBox1.SetItemChecked(j, true);
                }
            }
            else
            {
                for (int j = 0; j < checkedListBox1.Items.Count; j++)
                {
                    checkedListBox1.SetItemChecked(j, false);
                }
            }
        }

    }
}
