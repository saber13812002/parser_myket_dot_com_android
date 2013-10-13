using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
//using System.Drawing;
using System.IO;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;
//using System.Xml;
//using System.Xml.Linq;
using HtmlAgilityPack;
using System.Text;



namespace WindowsFormsApplication2
{


    public partial class Form1 : Form
    {
        DataTable dt_export_myket = new DataTable("myket");
        //Dictionary<string, int> dic = new Dictionary<string, int>();
        string path = @"C:\My Web Sites\cafebazar\http___myket.ir_\myket.ir";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt_export_myket.Columns.Add(new DataColumn("rate", typeof(float)));
            dt_export_myket.Columns.Add(new DataColumn("name", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("ver", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("lastupdatedatepersian", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("download", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("size", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("price", typeof(float)));
            dt_export_myket.Columns.Add(new DataColumn("creator", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("packagename", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("describtion", typeof(string)));
            dt_export_myket.Columns.Add(new DataColumn("similarapp", typeof(string)));

            textBox1.Text = "25";
            readAllFilesinSubfolders(int.Parse(textBox1.Text));
            //parse2222222222222("http://www.myket.ir/Default.aspx");


            //foreach (KeyValuePair<string, int> pair in dic)
            //{
            //    //Console.WriteLine("{0}, {1}",pair.Key,pair.Value);

            //    //if (lastpackagenam != packagename)
            //    //    dt_export_myket.Rows.Add(4.5, "name", 2, "1392/12/12", 5000, 5.9, 0, "creator", packagename, "describtion", "similarapp");
            //    dt_export_myket.Rows.Add(4.5, "name", pair.Value, "1392/12/12", 5000, 5.9, 0, "creator", pair.Key, "describtion", "similarapp");
            //}
        }

        private void readAllFilesinSubfolders(int count_of_)
        {
            string lastpackagename = "";

            var dir = new DirectoryInfo(path);
            int j = 1;
            foreach (var file in dir.EnumerateFiles("appd*.html", SearchOption.AllDirectories))
            {
                
                //if(file.ToString().Substring(0,8)=="appdetail")
                lastpackagename=parse(path + "\\" + file.ToString(), lastpackagename);
                if (j++ > count_of_)
                    break;
            }
        }

        private string parse(string fullPathoffile)
        {
            var hw = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc = hw.Load(fullPathoffile);
            //////////////////////////////////////////////////////<li class="cat_li"><a href="appdetail49c3.html?id=com.google.android.apps.maps"><img src="http://market.mservices.ir/myket/iconhandler.aspx?app=com.google.android.apps.maps&amp;width=80&amp;height=80" alt="" height="80px" width="80px"></a><ul><li><span>نام :</span> <a href="appdetail49c3.html?id=com.google.android.apps.maps"><b>نقشه گوگل</b></a></li><li><span>قیمت(تومان) :</span>رایگان</li><li><span>سازنده :</span><a style="color:blue" href="Applications03d8.html?type=Related&amp;Packagename=com.google.android.apps.maps">Google Inc.</a></li><li><div><img style="border-width:0px;" src="resources/RateStars/Rate45.png" alt="" height="17px" width="85px"></div></li></ul></li>
            //HtmlNode embed_tag = doc.DocumentNode.SelectSingleNode("//div[@class='c_trailer']");
            HtmlNode embed_tag = doc.DocumentNode.SelectSingleNode("//li[@class='cat_li']");

            HtmlNode embed;
            if (embed_tag.SelectSingleNode("embed[@id='template']") != null)
            {
                embed = embed_tag.SelectSingleNode("embed[@id='template']");

                return embed.InnerText;
            }
            else return "";
        }


        private string parse(string fullPathoffile,string lastpackagenam)
        {
            float rate = 0;
            string name = "";
            string ver="0";
            string lastupdatedatepersian = "";
            string download="0";
            string size="0MB";
            float price=0;
            string creator="";
            string packagename="";

            string txt=null;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(fullPathoffile);
            HtmlNode node = doc.DocumentNode;
            HtmlNodeCollection pre = node.SelectNodes("//div[@class='single_description']"); 
            //var prenodes = doc.DocumentNode.SelectNodes("//pre");
            if (pre != null)
            {
                foreach (var p in pre)
                {
                    HtmlNodeCollection a = p.SelectNodes("//ul[@class='single_info']");
                    foreach (var item in a)
                    {
                        HtmlNode embed = item.SelectSingleNode("//a[@id='Body_detailsControl_creatorApps']");
                        //HtmlNode stringrate = item.SelectSingleNode("//img[@id='Body_detailsControl_ratingDiv']");
                        if (embed != null)
                        {
                            packagename = findTag("Packagename", packagename, embed);
                            rate = findtagrate("rs/Rate", rate, item);
                            //rate daghigh
                            // ظ…غŒط§ظ†ع¯غŒظ† 4.34 ط§ط² 5
                            name = embed.InnerHtml;
                            //<span>ظ†ط³ط®ظ‡ :</span>
                            ver = findtagver("<span>ظ†ط³ط®ظ‡ :</span>","<", ver, item);
                            size = findtagsize("</li><li><span>ط³ط§غŒط² :</span>", "<", size, item);
                            lastupdatedatepersian = findtagdate("<span>ط¢ط®ط±غŒظ† ط¨ظ‡ ط±ظˆط²ط±ط³ط§ظ†غŒ :</span>", "<", lastupdatedatepersian, item);
                            download = findtagdl("<span>طھط¹ط¯ط§ط¯ ط¯ط§ظ†ظ„ظˆط¯ :</span>", "<", download, item);
                        }
                    }
                    if (lastpackagenam != packagename)
                        dt_export_myket.Rows.Add(rate, name, ver, lastupdatedatepersian, download, size, 0, "creator", packagename, "describtion", "similarapp");
                }
            }
            return packagename;
        }

        private string findtagdl(string thisstring, string tothistag, string download, HtmlNode item)
        {
            string dlssss = "0";
            string dlss = "0";
            string ht_cola = item.InnerHtml.ToString();
            if (ht_cola.Contains(thisstring))
            {
                int start_index = ht_cola.IndexOf(thisstring) + thisstring.Length;
                //int end_index = ht_cola.IndexOf("</li><li><span>ط³ط§غŒط² :</span>");
                //if (start_index < end_index)
                dlssss = ht_cola.Substring(start_index, 30);
                int endi = dlssss.IndexOf(tothistag);
                dlss = dlssss.Substring(0, endi-1);

                if (dlss.IndexOf(" ") > 0 && dlss.IndexOf(" ") < 6)
                {
                    endi = dlssss.IndexOf(" ");
                    dlss = dlss.Substring(0, endi);
                }
                //else if (dlss.IndexOf("²") > 0)
                //{
                //    int sti = dlssss.IndexOf("²");
                //    endi = dlssss.IndexOf(" ");
                //    dlss = dlss.Substring(13, endi - 13);
                //}
                //else if (dlss.IndexOf("ع") < 1)
                //{
                //    int sti = dlssss.IndexOf(" ");
                //    endi = dlssss.IndexOf("ع©ظ…طھط± ط§ط²")-"ع©ظ…طھط± ط§ط²".Length - sti;
                //    dlss = dlss.Substring(sti, endi);                //"rate","name","ver","lastupdatedatepersian","download","size","price","creator","packagename"
                //}
            }
            //if (int.Parse(dlss)!=null)
            //    download = int.Parse(dlss);
            ////else
            ////{
            ////    int endi = dlssss.IndexOf(" ");
            ////    dlss = dlss.Substring(0, endi);
            ////    download = int.Parse(dlss);
            ////}
            ////else
            ////    ver = float.Parse(vers[0].ToString());
            //return download;
            return dlss;
        }

        private string findtagdate(string thisstring, string tothistag, string lastupdatedatepersian, HtmlNode item)
        {
            string datessss = "1/11/2011";
            string datess = "1/11/2011";
            string ht_cola = item.InnerHtml.ToString();
            if (ht_cola.Contains(thisstring))
            {
                int start_index = ht_cola.IndexOf(thisstring) + thisstring.Length;
                //int end_index = ht_cola.IndexOf("</li><li><span>ط³ط§غŒط² :</span>");
                //if (start_index < end_index)
                datessss = ht_cola.Substring(start_index, 30);
                int endi = datessss.IndexOf(tothistag);
                datess = datessss.Substring(0, endi);

                //"rate","name","ver","lastupdatedatepersian","download","size","price","creator","packagename"
            }
            //if (float.Parse(vers) != 0)
            //size = float.Parse(sizess);
            //else
            //    ver = float.Parse(vers[0].ToString());
            return datess;
        }

        private string findtagsize(string thisstring, string tothistag, string size, HtmlNode item)
        {
            string sizessss = "0";
            string sizess = "0";
            string ht_cola = item.InnerHtml.ToString();
            if (ht_cola.Contains(thisstring))
            {
                int start_index = ht_cola.IndexOf(thisstring) + thisstring.Length;
                //int end_index = ht_cola.IndexOf("</li><li><span>ط³ط§غŒط² :</span>");
                //if (start_index < end_index)
                sizessss = ht_cola.Substring(start_index, 30);
                int endi = sizessss.IndexOf(tothistag);
                sizess = sizessss.Substring(0, endi);

                //"rate","name","ver","lastupdatedatepersian","download","size","price","creator","packagename"
            }
            //if (float.Parse(vers) != 0)
            //size = float.Parse(sizess);
            //else
            //    ver = float.Parse(vers[0].ToString());
            return sizess;
        }

        private string findtagver(string thisstring,string tothistag, string ver, HtmlNode item)
        {
            string veresss = "0";
            string vers = "0";
            string ht_cola = item.InnerHtml.ToString();
            if (ht_cola.Contains(thisstring))
            {
                int start_index = ht_cola.IndexOf(thisstring) + thisstring.Length;
                //int end_index = ht_cola.IndexOf("</li><li><span>ط³ط§غŒط² :</span>");
                //if (start_index < end_index)
                    veresss = ht_cola.Substring(start_index, 30);
                    int endi = veresss.IndexOf(tothistag);
                vers = veresss.Substring(0, endi);

                //"rate","name","ver","lastupdatedatepersian","download","size","price","creator","packagename"
            }
            //if (float.Parse(vers) != 0)
            //    ver = float.Parse(vers);
            //else
            //    ver = float.Parse(vers[0].ToString());
            return vers;
        }

        private float findtagrate(string thisstring, float rate, HtmlNode item)
        {
            string ratesss="0";
            string ht_cola = item.InnerHtml.ToString();
            if (ht_cola.Contains(thisstring))
            {
                int start_index = ht_cola.IndexOf(thisstring) + thisstring.Length;
                int end_index = ht_cola.IndexOf(".png");
                if (start_index < end_index)
                    ratesss = ht_cola.Substring(start_index, end_index - start_index);
                //"rate","name","ver","lastupdatedatepersian","download","size","price","creator","packagename"
            }
            if (float.Parse(ratesss) != 0)
                rate = float.Parse(ratesss) / 10;
            else
                rate = float.Parse(ratesss[0].ToString());
            return rate;
        }

        private static string findTag(string thisstring,string packagename, HtmlNode embed)
        {
            //string packagename ="";
            string ht_cola = embed.OuterHtml.ToString();
            if (ht_cola.Contains(thisstring))
            {
                int start_index = ht_cola.IndexOf(thisstring) + thisstring.Length;
                int end_index = ht_cola.IndexOf("style=");
                if (start_index < end_index)
                    packagename = ht_cola.Substring(start_index, end_index - start_index - 2);
                //"rate","name","ver","lastupdatedatepersian","download","size","price","creator","packagename"
            }
            return packagename;
        }
        
        private string read(string fileName)
        {
            Encoding encoding = Encoding.GetEncoding(1252);
            string text = File.ReadAllText(fileName, encoding);

            return text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.ResetText();

            HtmlAgilityPack.HtmlWeb w = new HtmlAgilityPack.HtmlWeb();
            HtmlAgilityPack.HtmlDocument doc = w.Load("http://www.farsnews.com");
            HtmlAgilityPack.HtmlNodeCollection nc = doc.DocumentNode.SelectNodes(".//div[@class='topnewsinfotitle']");

            listBox1.Items.Add(nc.Count + " Items selected!");

            foreach (HtmlAgilityPack.HtmlNode node in nc)
            {
                listBox1.Items.Add(node.InnerText);
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            if (dt_export_myket.Rows.Count > 0)
                export_to_excel_tb2(dt_export_myket,1);
            else
                textBox1.Text = "0";
        }

        private void export_to_excel_tb2(DataTable data,int a)
        {
            var dia = new System.Windows.Forms.SaveFileDialog();
            dia.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dia.Filter = "Excel Worksheets (*.xlsx)|*.xlsx|xls file (*.xls)|*.xls|All files (*.*)|*.*";
            if (dia.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                //data = null;// use the DataSource of the DataGridView here
                var excel = new OfficeOpenXml.ExcelPackage();
                var ws = excel.Workbook.Worksheets.Add("worksheet-name");
                // you can also use LoadFromCollection with an `IEnumerable<SomeType>`
                ws.Cells["A1"].LoadFromDataTable(data, true, OfficeOpenXml.Table.TableStyles.Light1);
                ws.Cells[ws.Dimension.Address.ToString()].AutoFitColumns();

                using (var file = File.Create(dia.FileName))
                    excel.SaveAs(file);
            }
        }

        private void btnReplay_Click(object sender, EventArgs e)
        {
            readAllFilesinSubfolders(int.Parse(textBox1.Text));
        }

        //com.monotype.android.font.pack.retro
    }
}
