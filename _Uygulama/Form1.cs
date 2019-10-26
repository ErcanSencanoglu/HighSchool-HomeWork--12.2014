using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using HtmlAgilityPack;
using System.IO;
using System.Threading;
using Microsoft.Office.Interop.Excel;

namespace _Uygulama
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public void firmalariAl()
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.GetEncoding("UTF-8");

            List<string> firmalar = new List<string>();
            Uri firmaUrl = new Uri("https://www.sikayetvar.com/firmalar");
            string html = client.DownloadString(firmaUrl);
            HtmlAgilityPack.HtmlDocument htmldoc = new HtmlAgilityPack.HtmlDocument();
            htmldoc.LoadHtml(html);



            HtmlNodeCollection firmaSayisi = htmldoc.DocumentNode.SelectNodes("//div[@class='pager']/a");
            int son = Convert.ToInt32(firmaSayisi[firmaSayisi.Count - 2].Attributes["title"].Value);
            //int son = Convert.ToInt32(firmalar[firmalar.Count - 1].InnerText;
            for (int i = 0; i < son; i++)
            {
                firmaUrl = new Uri("https://www.sikayetvar.com/firmalar/p/" + i + "#pager");
                html = client.DownloadString(firmaUrl);
                htmldoc.LoadHtml(html);
                HtmlNodeCollection firmalar2 = htmldoc.DocumentNode.SelectNodes("//ul[@class='firmalar kategoriFirmalar']/li/a");
                foreach (var baslik in firmalar2) { comboBox1.Items.Add(baslik.Attributes["title"].Value); }
            }
            listBox1.Items.Add(comboBox1.Items.Count + " adet firma bulundu. " + DateTime.Now);
            listBox1.SelectedIndex = listBox1.Items.Count - 1;

        }

        public int sonAl(string seciliFirma) {
        
            int son =0;
            WebClient client = new WebClient();
            client.Encoding = Encoding.GetEncoding("UTF-8");

            List<string> firmalar = new List<string>();
            Uri firmaUrl = new Uri("https://www.sikayetvar.com/"+seciliFirma);
            string html = client.DownloadString(firmaUrl);
            HtmlAgilityPack.HtmlDocument htmldoc = new HtmlAgilityPack.HtmlDocument();
            htmldoc.LoadHtml(html);

            HtmlNodeCollection firmaSayisi = htmldoc.DocumentNode.SelectNodes("//ul[@class='sikayetInfoNav']/li/a/p");
            son = Convert.ToInt32(firmaSayisi[0].InnerText)/10;
            if (Convert.ToInt32(firmaSayisi[0].InnerText) % 2 == 0 && son !=0) { son--; }

            return son;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            firmalariAl();
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox1.Sorted = true;
            listView1.View = View.Details;
            listView1.Columns.Add("Fima", 100, HorizontalAlignment.Left);
            listView1.Columns.Add("Yıl", 100, HorizontalAlignment.Left);
            listView1.Columns.Add("Durum", 100, HorizontalAlignment.Left);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listView1.Items.Count == 0) { MessageBox.Show("Listeye firma eklenmemiş"); return; }
            String ornekDosya = ornekDosyaSec();
            String kayıtYol = kayitYoluSec();

            for (int f = 0; f < listView1.Items.Count; f++)
            {
                listView1.Items[f].SubItems[2].Text = "Devam Ediyor.";
                listView1.Items[f].BackColor = Color.Olive;
            
            string seciliFirma = "";
            string firmaismi = listView1.Items[f].Text.ToString().ToLower();
            seciliFirma = listView1.Items[f].Text.ToString().ToLower();
         //   seciliFirma = comboBox1.SelectedItem.ToString().ToLower();
            seciliFirma =  seciliFirma.Replace(" ", "-");
            seciliFirma = seciliFirma.Replace('ı', 'i');
            seciliFirma = seciliFirma.Replace('ü', 'u');
            seciliFirma = seciliFirma.Replace('ö', 'o');
            seciliFirma = seciliFirma.Replace(".", "");
            seciliFirma = seciliFirma.Replace('ş', 's');
            seciliFirma = seciliFirma.Replace('ç', 'c');


            string seciliTarih = listView1.Items[f].SubItems[0].Tag.ToString();

            listBox1.Items.Add("Uygulama çalışmaya başladı: " + DateTime.Now);
            listBox1.SelectedIndex = listBox1.Items.Count - 1;

            ExcelUygulama = new Microsoft.Office.Interop.Excel.Application();
            ExcelUygulama.Visible = true;

            WebClient client = new WebClient();
            client.Encoding = Encoding.GetEncoding("UTF-8");

         




            int son = sonAl(seciliFirma);
            for (int i = 0; i <= son; i++)//(int i = baslangıc; i >= bitis; i--)
            {

                listBox1.Items.Add(i + "/" + son + " sayfadaki şikayetler alınmaya başlandı.");
                listBox1.SelectedIndex = listBox1.Items.Count - 1;

                Uri url = new Uri("https://www.sikayetvar.com/"+seciliFirma+"?page=" + "0" + "&dateRange="+seciliTarih);
                webBrowser1.Navigate(url);
                listBox1.Items.Add("Browser ın yüklemesi bekleniyor"); listBox1.ForeColor = Color.Red;
                listBox1.SelectedIndex = listBox1.Items.Count - 1;
                while (webBrowser1.ReadyState != WebBrowserReadyState.Complete)
                { System.Windows.Forms.Application.DoEvents(); }
                listBox1.Items.Add("Browser yüklendi"); listBox1.ForeColor = Color.Green;
                listBox1.SelectedIndex = listBox1.Items.Count - 1;

                //string html = client.DownloadString(url);
                HtmlAgilityPack.HtmlDocument htmldoc = new HtmlAgilityPack.HtmlDocument();
                htmldoc.Load(webBrowser1.DocumentStream);
                HtmlNodeCollection basliklar = htmldoc.DocumentNode.SelectNodes("//li[@class='sikayetDurumBilgisi']/a");
                listBox1.Items.Add(i + "/" + son + " sayfadaki şikayetler alındı.\nAlınan toplam Başlık: " + basliklar.Count);
                listBox1.SelectedIndex = listBox1.Items.Count - 1;

                List<string> liste = new List<string>();
                foreach (var baslik in basliklar) { liste.Add(baslik.Attributes["href"].Value); }
                listBox1.Items.Add(i + "/" + son + " sayfadaki şikayetler listeye eklendi.\nListedeki şikayet sayısı: " + liste.Count);
                listBox1.SelectedIndex = listBox1.Items.Count - 1;
                //Teşekkürler silinecek ve 2013 verileri alınacak
                //liste.Reverse();


                listBox1.Items.Add("Listedeki veriler alınmaya başlandı.");
                listBox1.SelectedIndex = listBox1.Items.Count - 1;
                for (int j = 0; j < liste.Count; j++)
                {
                    List<string> bilgiler = new List<string>();

                    Uri bilGiUrl = new Uri("https://www.sikayetvar.com/" + liste[j]);
                    string html2 = client.DownloadString(bilGiUrl);
                    HtmlAgilityPack.HtmlDocument htmldoc2 = new HtmlAgilityPack.HtmlDocument();
                    htmldoc2.LoadHtml(html2);
                    listBox1.Items.Add((j + 1) + ".Şikayet verileri alınmaya başladı.");
                    listBox1.SelectedIndex = listBox1.Items.Count - 1;

                    //Şikayet no
                    HtmlNodeCollection basliklar2 = htmldoc2.DocumentNode.SelectNodes("//div[@class='sikayetDetayNavigasyon sp']/b");
                    bilgiler.Add((basliklar2[0].InnerText.Substring(13).TrimStart().TrimEnd() != string.Empty)
                        ? basliklar2[0].InnerText.Substring(13).TrimStart().TrimEnd() : " ");

                    //Şikayet başlık

                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//div[@class='sikayetBaslik']/div/h1");
                    string b = basliklar2[0].InnerText;
                    bilgiler.Add((basliklar2[0].InnerText.Trim().Substring(6).TrimStart().TrimEnd() != string.Empty)
                        ? basliklar2[0].InnerText.Trim().Substring(6).TrimStart().TrimEnd() : " " );


                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//div[@class='sikayetBaslik']/div/h1");
                    bilgiler.Add((basliklar2[0].InnerText != string.Empty) ? basliklar2[0].InnerText : " ");

                    //Üye Linki
                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//div[@class='sikayetBaslik']/div/span/a");
                    Uri uyeBilgi = new Uri("https://www.sikayetvar.com/" + basliklar2[0].Attributes["href"].Value);


                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//div[@class='sikayetBaslik']/div/span");
                    bilgiler.Add((basliklar2[0].InnerText.Substring(basliklar2[0].InnerText.IndexOf('|') + 2) != string.Empty)
                        ? basliklar2[0].InnerText.Substring(basliklar2[0].InnerText.IndexOf('|') + 2) : " ");


                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//p[@class='sikayetDetayMetin']");
                    if (basliklar2 == null) { continue; }
                    bilgiler.Add((basliklar2[0].InnerText.TrimStart().TrimEnd() != string.Empty) ? basliklar2[0].InnerText.TrimStart().TrimEnd() : " ");

                    //Şikayet Konuları
                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//ul[@class='konuEtiketleri']/li/a");
                    List<string> sikayetKonulari = new List<string>();
                    if (basliklar2 != null)
                    {
                        int sikayatSayisi = (basliklar2.Count < 4) ? basliklar2.Count : 3;
                        for (int artis = 0; artis < sikayatSayisi; artis++)
                        { sikayetKonulari.Add(basliklar2[artis].InnerText); }
                    }

                    //Şikayet Cevapları
                    basliklar2 = htmldoc2.DocumentNode.SelectNodes("//ul[@class='sikayetCevaplar']/li/div/div[@class='sikayetCevapMetin']/p");
                    string metin = "";
                    if (basliklar2 != null) { foreach (var baslik in basliklar2) { metin = metin + baslik.InnerText; metin += (basliklar2.Count > 1) ? "\n-------------\n" : ""; } } bilgiler.Add(metin);

                    listBox1.Items.Add((j + 1) + ".Şikayet verileri  alındı.");
                    listBox1.SelectedIndex = listBox1.Items.Count - 1;
                    listBox1.Items.Add((j + 1) + ".Üyelik verileri alınmaya başladı.");
                    listBox1.SelectedIndex = listBox1.Items.Count - 1;

                    html2 = client.DownloadString(uyeBilgi);
                    htmldoc2 = new HtmlAgilityPack.HtmlDocument();
                    htmldoc2.LoadHtml(html2);


                    //Üye ismi
                    HtmlNodeCollection basliklar3 = htmldoc2.DocumentNode.SelectNodes("//div[@class='uyeDetayTitle']/h2[@class='sp']");
                    if (basliklar3 == null) { continue; }
                    bilgiler.Add((basliklar3[0].InnerText.TrimStart().TrimEnd() != string.Empty) ? basliklar3[0].InnerText.TrimStart().TrimEnd() : " ");


                    //ÜyelikTarihi
                    basliklar3 = htmldoc2.DocumentNode.SelectNodes("//div[@class='uyeDetayBilgiler']/div/div/div");
                    bilgiler.Add((basliklar3[0].InnerText != string.Empty) ?
                        basliklar3[0].InnerText.TrimStart().TrimEnd().Substring(
                        basliklar3[0].InnerText.IndexOf(':'), basliklar3[0].InnerText.IndexOf("Üyelik Seviyesi") - basliklar3[0].InnerText.IndexOf(':') - 1)
                        .TrimStart().TrimEnd() : " ");

                    //Şikayet Sayısı
                    bilgiler.Add((basliklar3[1].InnerText.Substring(basliklar3[1].InnerText.IndexOf(':'),
                                    basliklar3[1].InnerText.LastIndexOf("Yorum Sayısı ") - liste[1].IndexOf(':') - 1).Trim())
                                     != string.Empty
                                     ? basliklar3[1].InnerText.Substring(
                                     basliklar3[1].InnerText.IndexOf(':') + 2,
                                   basliklar3[1].InnerText.Length - basliklar3[1].InnerText.LastIndexOf("Yorum Sayısı ") - basliklar3[1].InnerText.IndexOf(':') - 1).Trim()
                                    : " "
                        );

                    //Yorum Sayısı
                    bilgiler.Add((basliklar3[1].InnerText.Substring(basliklar3[1].InnerText.LastIndexOf(':') + 2).TrimStart().TrimEnd() != string.Empty)
                        ? basliklar3[1].InnerText.Substring(basliklar3[1].InnerText.LastIndexOf(':') + 2).TrimStart().TrimEnd() : " ");

                    listBox1.Items.Add((j + 1) + ".Veri alındı.");
                    listBox1.SelectedIndex = listBox1.Items.Count - 1;



                    string tarih = bilgiler[3].Substring(3, bilgiler[3].Length - 8).TrimEnd();
                    string simdikiYol = kayıtYol + "/" + seciliFirma+" "+tarih + ".xlsx";
                    // string simdikiYol = "C:\\Users/Ercan/Desktop/Yazılımlar/ÖdevDenemesi/dosyalar/" + tarih + ".xlsx";
                    if (oncekiYol == simdikiYol && oncekiYol != string.Empty)
                    {
                        listBox1.Items.Add("Kayıtlar şuanki sayfaya eklendi.");
                        listBox1.SelectedIndex = listBox1.Items.Count - 1;
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 2];
                        alan.Value2 = bilgiler[0];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 3];
                        alan.Value2 = string.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(bilgiler[3]));

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 4];
                        alan.Value2 = bilgiler[2];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 5];
                        alan.Value2 = bilgiler[6];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 6];
                        alan.Value2 = string.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(bilgiler[7]));

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 7];
                        alan.Value2 = bilgiler[8];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 8];
                        alan.Value2 = bilgiler[9];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 9];
                        alan.Value2 = bilgiler[4];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 13];
                        alan.Value2 = bilgiler[1];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 14];
                        alan.Value2 = bilgiler[5];
                        for (int n = 10; n < 10 + sikayetKonulari.Count; n++)
                        {
                            alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, n];
                            alan.Value2 = sikayetKonulari[n - 10];
                        }
                        kacinciSatir++;
                        listBox1.Items.Add("Kayıtlar şuanki sayfaya eklendi. Yeni satır sayısı:" + kacinciSatir);
                        listBox1.SelectedIndex = listBox1.Items.Count - 1;
                    }//if sonu
                    else if (oncekiYol == string.Empty)
                    {
                        oncekiYol = simdikiYol;
                        if (System.IO.File.Exists(simdikiYol))
                        {
                            CalismaKitabi = ExcelUygulama.Workbooks.Open(simdikiYol);
                            listBox1.Items.Add("Açık bir sayfa olmadığı için yeni bir sayfa açıldı. Sayfa adı:" + tarih);
                            listBox1.SelectedIndex = listBox1.Items.Count - 1;
                            for (int m = 3; m < CaslismaSayfasi.Rows.Count; m++)
                            {
                                alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[m, 2];
                                if (alan.get_Value() == null) { kacinciSatir = m; break; }
                            }

                        }
                        else
                        {
                            CalismaKitabi = ExcelUygulama.Workbooks.Open(ornekDosya);
                            //CalismaKitabi = ExcelUygulama.Workbooks.Open("C:\\Users/Ercan/Desktop/Yazılımlar/ÖdevDenemesi/dosyalar/ornek.xlsx");
                            kaydet = true;
                            listBox1.Items.Add("Açık bir sayfa olmadığı için yeni bir sayfa oluşturuldu. Sayfa adı:" + tarih);
                            listBox1.SelectedIndex = listBox1.Items.Count - 1;
                        }

                        CaslismaSayfasi = (Microsoft.Office.Interop.Excel.Worksheet)CalismaKitabi.Sheets[1];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 2];
                        alan.Value2 = bilgiler[0];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 3];
                        alan.Value2 = string.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(bilgiler[3]));

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 4];
                        alan.Value2 = bilgiler[2];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 5];
                        alan.Value2 = bilgiler[6];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 6];
                        alan.Value2 = string.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(bilgiler[7]));

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 7];
                        alan.Value2 = bilgiler[8];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 8];
                        alan.Value2 = bilgiler[9];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 9];
                        alan.Value2 = bilgiler[4];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 13];
                        alan.Value2 = bilgiler[1];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 14];
                        alan.Value2 = bilgiler[5];
                        for (int n = 10; n < 10 + sikayetKonulari.Count; n++)
                        {
                            alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, n];
                            alan.Value2 = sikayetKonulari[n - 10];
                        }
                        kacinciSatir++;
                        listBox1.Items.Add("Kayıtlar şuanki sayfaya eklendi. Yeni satır sayısı:" + kacinciSatir);
                        listBox1.SelectedIndex = listBox1.Items.Count - 1;

                    }//else if sonu
                    else
                    {
                        if (kaydet == false)
                        {
                            CalismaKitabi.Save();
                            CalismaKitabi.Close();
                            kaydet = false;
                            listBox1.Items.Add("Şuanki sayfa kaydedilip kapandı.");
                            listBox1.SelectedIndex = listBox1.Items.Count - 1;
                        }
                        else
                        {
                            CalismaKitabi.SaveAs(@oncekiYol,
                            Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                            Type.Missing,
                            Type.Missing,
                            false,
                            Type.Missing,
                            Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive);
                            CalismaKitabi.Close();
                            kaydet = false;
                            listBox1.Items.Add("Şuanki sayfa farklı kaydedilip kapandı.");
                            listBox1.SelectedIndex = listBox1.Items.Count - 1;
                        }

                        oncekiYol = simdikiYol;
                        if (System.IO.File.Exists(simdikiYol))
                        {
                            CalismaKitabi = ExcelUygulama.Workbooks.Open(simdikiYol);
                            listBox1.Items.Add("Açık bir sayfa olmadığı için yeni bir sayfa açıldı. Sayfa adı:" + tarih);
                            listBox1.SelectedIndex = listBox1.Items.Count - 1;
                        }
                        else
                        {
                            //CalismaKitabi = ExcelUygulama.Workbooks.Open("C:\\Users/Ercan/Desktop/Yazılımlar/ÖdevDenemesi/dosyalar/ornek.xlsx");
                            CalismaKitabi = ExcelUygulama.Workbooks.Open(ornekDosya);
                            kaydet = true;
                            listBox1.Items.Add("Açık bir sayfa olmadığı için yeni bir sayfa oluşturuldu. Sayfa adı:" + tarih);
                            listBox1.SelectedIndex = listBox1.Items.Count - 1;
                        }

                        CaslismaSayfasi = (Microsoft.Office.Interop.Excel.Worksheet)CalismaKitabi.Sheets[1];
                        for (int m = 3; m < CaslismaSayfasi.Rows.Count; m++)
                        {
                            alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[m, 2];
                            if (alan.get_Value() == null) { kacinciSatir = m; break; }
                        }
                        listBox1.Items.Add("Yeni sayfadaki satır sayısı belirlendi:" + kacinciSatir);
                        listBox1.SelectedIndex = listBox1.Items.Count - 1;

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 2];
                        alan.Value2 = bilgiler[0];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 3];
                        alan.Value2 = string.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(bilgiler[3]));

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 4];
                        alan.Value2 = bilgiler[2];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 5];
                        alan.Value2 = bilgiler[6];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 6];
                        alan.Value2 = string.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(bilgiler[7]));

                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 7];
                        alan.Value2 = bilgiler[8];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 8];
                        alan.Value2 = bilgiler[9];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 9];
                        alan.Value2 = bilgiler[4];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 13];
                        alan.Value2 = bilgiler[1];
                        alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, 14];
                        alan.Value2 = bilgiler[5];
                        for (int n = 10; n < 10 + sikayetKonulari.Count; n++)
                        {
                            alan = (Microsoft.Office.Interop.Excel.Range)CaslismaSayfasi.Cells[kacinciSatir, n];
                            alan.Value2 = sikayetKonulari[n - 10];
                        }
                        kacinciSatir++;
                        listBox1.Items.Add("Kayıtlar şuanki sayfaya eklendi. Yeni satır sayısı:" + kacinciSatir);
                    }//else sonu

                    listBox1.SelectedIndex = listBox1.Items.Count - 1;


                }//2.for sonu
            }//for sonu

            if (kaydet == false)
            {
                CalismaKitabi.Save();
                CalismaKitabi.Close();
                kaydet = false;
                listBox1.Items.Add("Şuanki sayfa kaydedilip kapandı.");
                listBox1.SelectedIndex = listBox1.Items.Count - 1;
            }
            else
            {
                CalismaKitabi.SaveAs(@oncekiYol,
                Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                Type.Missing,
                Type.Missing,
                false,
                Type.Missing,
                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive);
                CalismaKitabi.Close();
                kaydet = false;
                listBox1.Items.Add("Şuanki sayfa farklı kaydedilip kapandı.");
                listBox1.SelectedIndex = listBox1.Items.Count - 1;
            }

            oncekiYol = "";
            ExcelUygulama.Quit();
            listView1.Items[f].SubItems[2].Text = "Tamamlandı.";
            listView1.Items[f].BackColor = Color.Green;
            }
            listBox1.Items.Add("Uygulama sonlandı: " + DateTime.Now);
            listBox1.ForeColor = Color.Red;
            listBox1.SelectedIndex = listBox1.Items.Count - 1;

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Return)
            {
                if (comboBox1.Items.IndexOf(textBox1.Text) != -1)
                {
                    comboBox1.SelectedItem = textBox1.Text;
                }
                else { MessageBox.Show("Firma Bulunamadı.", "Bilgi"); }
            }
        }

        string oncekiYol = "";
        int kacinciSatir = 3;
        bool kaydet = false;
        Microsoft.Office.Interop.Excel.Application ExcelUygulama;
        Microsoft.Office.Interop.Excel.Workbook CalismaKitabi;
        Microsoft.Office.Interop.Excel.Worksheet CaslismaSayfasi;
        Microsoft.Office.Interop.Excel.Range alan;

        private String ornekDosyaSec()
        {
            OpenFileDialog ornekDosyaDialog = new OpenFileDialog();
            ornekDosyaDialog.Title = "Örnek Bir Excel Dosyası Seçin";
            ornekDosyaDialog.InitialDirectory = "C:\\";
            ornekDosyaDialog.Filter = "Excel Dosyası|*.xlsx";
            while (!(ornekDosyaDialog.ShowDialog() == DialogResult.OK)) { MessageBox.Show("Bir tane örnek excel dosyası seçilmeli.", "Uyarı"); }
            return ornekDosyaDialog.FileName;
        }

        private String kayitYoluSec()
        {
            FolderBrowserDialog ornekDosyaDialog = new FolderBrowserDialog();
            ornekDosyaDialog.ShowNewFolderButton = true;
            while (!(ornekDosyaDialog.ShowDialog() == DialogResult.OK)) { MessageBox.Show("Kayıt yapılacak yol seçilmeli", "Uyarı"); }
            return ornekDosyaDialog.SelectedPath;
        }

        private void btnEkle_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < listView1.Items.Count; i++)
            {
                if (listView1.Items[i].SubItems[0].Text.ToString() == comboBox1.SelectedItem.ToString() &&
                    listView1.Items[i].SubItems[1].Text.ToString() == comboBox2.SelectedItem.ToString()) {
                        MessageBox.Show("Bu kayıt daha önce eklenmiş"); return;
                }
            }
            if (comboBox1.SelectedIndex != -1 && comboBox2.SelectedIndex != -1)
            {
                ListViewItem item = new ListViewItem();
                listView1.Items.Add(comboBox1.SelectedItem.ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add(comboBox2.SelectedItem.ToString());
                listView1.Items[listView1.Items.Count - 1].SubItems.Add("Başlamadı");
                listView1.Items[listView1.Items.Count - 1].BackColor = Color.Red;

                string seciliTarih = "L1Y";
                if (comboBox2.SelectedIndex == 0) { seciliTarih = "-1"; }
                else if (comboBox2.SelectedIndex == 1) { seciliTarih = "L1Y"; }
                else if (comboBox2.SelectedIndex == 2) { seciliTarih = "Y13"; }
                else if (comboBox2.SelectedIndex == 3) { seciliTarih = "L14"; }
                else if (comboBox2.SelectedIndex == 4) { seciliTarih = "L15"; }
                listView1.Items[listView1.Items.Count - 1].SubItems[0].Tag = seciliTarih;
               
            }
            else { MessageBox.Show("Bir firma seçiniz.","Bilgi"); }
        }

      
    }
}
