using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Speech.Recognition;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;
using System.IO;
using System.Text.RegularExpressions;

namespace SpeechRecognition
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //static Form1 f1;
        static Boolean action = false;
        static Boolean power = true;
        static Boolean isweather = false;
        static Boolean isemail = false;
        static Boolean OpenFile = false;
        static Label l;
        static Label l2;
        static Label l3;
        static TextBox t1;

        static void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            //Boolean action = false;
            //Boolean power = true;
            string txt = e.Result.Text;
            //l3.Text = txt;
            t1.Text = t1.Text + " " + Form1.isweather.ToString();
            Clipboard.SetText(txt);
            SendKeys.SendWait("^{v}");
            if (e.Result.Confidence < 0.7) return;
            //if (power == true)
            if (Form1.power == true)
            {
                if (e.Result.Confidence > 0.7) l.Text = l.Text + " " + e.Result.Text + "; ";
                //if (txt.IndexOf("новая вкладка") >= 0)
                if (!Form1.action & !Form1.isweather & !Form1.isemail)
                {
                    switch (e.Result.Text.ToString())
                    {
                        case "привет":
                            MessageBox.Show("Здравствуйте");
                            break;
                        case "Пока":
                            MessageBox.Show("До свидания");
                            break;
                        case "открыть вкладку":
                            Process.Start("chrome.exe", "-new-tab https://www.yandex.ru/");
                            break;
                        case "открой проводник":
                            l.Text = "Вызываю компьютер";
                            Process.Start("explorer");
                            break;
                        case "выключить микрофон":
                            Form1.power = false;
                            l2.Text = "off";
                            break;
                        case "включить микрофон":
                            Form1.power = true;
                            l2.Text = "off";
                            break;
                        case "открыть":
                            Form1.action = true;
                            l3.Text = "open";
                            break;
                        case "погода":
                            MessageBox.Show("Москва +18");
                            Form1.isweather = true;
                            l3.Text = "weather";
                            break;
                        case "письмо":
                            Form1.isemail = true;
                            l3.Text = "email";
                            break;
                    }
                }
                else if (Form1.action)
                {
                    switch (e.Result.Text.ToString())
                    {
                        case "блокнот":
                            l.Text = "открываю блокнот";
                            Process.Start("notepad");
                            l3.Text = "close"; //
                            action = false;
                            break;
                        case "проводник":
                            l.Text = "Вызываю компьютер";
                            Process.Start("explorer");
                            l3.Text = "close"; //
                            action = false;
                            break;
                        case "калькулятор":
                            l.Text = "калькулятор";
                            Process.Start("calc");
                            action = false;
                            l3.Text = "close"; //
                            break;
                        case "файл":
                            l.Text = "открываю файл";
                            Form1.OpenFile = true;
                            break;
                    }
                }
                else if (Form1.isweather)
                {
                    //инициализируем город
                    Form1.isweather = false;
                    string city;
                    string id_city;
                    city = e.Result.Text.ToString();
                    var doc = XDocument.Load("https://pogoda.yandex.ru/static/cities.xml");
                    var cityId = doc.Root.Descendants("city").Select(n => new { city = n.Value, id = n.Attribute("id").Value });
                    //l2.Text = cityId.FirstOrDefault(n => n.city == city)?.id;
                    id_city = cityId.FirstOrDefault(n => n.city == city)?.id;

                    //получение погоды
                    var t = XDocument.Load(string.Format("http://export.yandex.ru/weather-ng/forecasts/{0}.xml", id_city));

                    XNamespace ya = "http://weather.yandex.ru/forecast";
                    var fact = t.Document.Root.Element(ya + "fact");

                    

                    //MessageBox.Show(fact.Element(ya + "station").Value + " " + fact.Element(ya + "temperature").Value_
                    //  + " градусов " + fact.Element(ya + "weather_type").Value);

                }
                else if (Form1.isemail)
                {
                    Form1.isemail = false;
                    //инициализируем получателя
                    string[] arr;
                    string result = e.Result.Text.ToString();
                    string recipient = null;
                    MessageBox.Show("Письмо " + result);
                    arr = File.ReadAllLines(@"D:\Misha\recognization\Files\Emails.txt");
                    for (int i = 0; i < arr.Length; i++)
                    {
                        if (result == arr[i].Substring(0, arr[i].IndexOf("|")))
                        {
                            recipient = arr[i].Substring(arr[i].IndexOf("|") + 1, arr[i].Length - arr[i].IndexOf("|") - 1);
                            break;
                        }
                    }

                    t1.Text = t1.Text + " resepient " + recipient;

                    if (recipient != null & recipient != "")
                    {
                        //отправляем письмо
                        l3.Text = "Отправляем в почту";
                        MailAddress from = new MailAddress("Hanter715@gmail.com", "Misha");
                        MailAddress to = new MailAddress(recipient);
                        MailMessage m = new MailMessage(from, to);
                        m.Subject = "Тест";
                        m.Body = "<h2>Письмо-тест работы smtp-клиента</h2>";
                        m.IsBodyHtml = true;
                        SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
                        smtp.Credentials = new NetworkCredential(from.Address, "9misha97");
                        smtp.EnableSsl = true;
                        smtp.Send(m);
                    }
                }
                else if (Form1.OpenFile)
                {
                    Regex file = new Regex(@".*\.txt$");
                    DirectoryInfo dr = new DirectoryInfo(@"E:\");
                    Search(dr, file);
                }
                    if (txt.IndexOf("погодs") >= 0)
                {
                    //l3.Text = "Получаем погоду";

                    //var doc = XDocument.Load("https://pogoda.yandex.ru/static/cities.xml");


                    //var cityId =
                        //doc.Root.Descendants("city")
                      //      .Select(n => new { city = n.Value, id = n.Attribute("id").Value });


                    //var city = "Москва";
                    //l3.Text = cityId.FirstOrDefault(n => n.city == city)?.id;
                    //27612

                    //string id_city = "27612";

                    //var t = XDocument.Load(string.Format("http://export.yandex.ru/weather-ng/forecasts/{0}.xml", id_city));
                    //var t = XDocument.Load("https://api.weather.yandex.ru/v1/forecast?geoid=213&l10n=true");

                   // XNamespace ya = "http://weather.yandex.ru/forecast";
                    //var fact = t.Document.Root.Element(ya + "fact");

                   // Console.WriteLine(fact.Element(ya + "station").Value);
                    //Console.WriteLine(fact.Element(ya + "temperature").Value + " градусов");
                   // Console.WriteLine(fact.Element(ya + "weather_type").Value);
                    //Console.ReadKey(true);
                }
              }

            else
            {
                return;
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            //загрузка основной формы/программы
            l = label1;
            l2 = label2;
            l3 = label3;
            t1 = textBox3;

            //инициализация языковой библиотеки и голосового движка
           System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("ru-ru");
            SpeechRecognitionEngine sre = new SpeechRecognitionEngine(ci);
            sre.SetInputToDefaultAudioDevice();

            //
            sre.EndSilenceTimeoutAmbiguous = TimeSpan.FromSeconds(0);
            sre.EndSilenceTimeout = TimeSpan.FromSeconds(0);
            sre.BabbleTimeout = TimeSpan.FromSeconds(0);
            sre.InitialSilenceTimeout = TimeSpan.FromSeconds(0);
            //

            //обьявления события обработчиа синтеза
            sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);

            String path = Directory.GetCurrentDirectory();

            //загрузка переменных в конструктор грамматики
            string[] resarr;
            int count=0;
            //resarr = File.ReadAllLines(@"D:\Misha\recognization\Files\Emails.txt");
            resarr = File.ReadAllLines(path+@"\Files\Emails.txt");
            foreach (string elstr in resarr)
            {
                resarr[count] = elstr.Substring(0, elstr.IndexOf("|"));
                count = count + 1;
                t1.Text = t1.Text+"Имя: " + elstr.Substring(0, elstr.IndexOf("|")) + "      Адрес: " + elstr.Substring(elstr.IndexOf("|") + 1, elstr.Length - elstr.IndexOf("|") - 1) + "\r\n";
            }

            //String path = Directory.GetCurrentDirectory();

            Choices choises = new Choices();
            //choises.Add(File.ReadAllLines(@"D:\Misha\recognization\Files\Command.txt"));
            //choises.Add(File.ReadAllLines(@"D:\Misha\recognization\Files\Numbers.txt"));
            //choises.Add(File.ReadAllLines(@"D:\Misha\recognization\Files\Cities.txt"));
            choises.Add(File.ReadAllLines(path+@"\Files\Command.txt"));
            choises.Add(File.ReadAllLines(path+@"\Files\Numbers.txt"));
            choises.Add(File.ReadAllLines(path+@"\Files\Cities.txt"));
            choises.Add(resarr);

            //инициализируем конструктор грамматики
            GrammarBuilder gb = new GrammarBuilder();
            gb.Append(choises);
            gb.Culture = ci;
            
            
            Grammar g = new Grammar(gb);
            sre.LoadGrammar(g);

            sre.RecognizeAsync(RecognizeMode.Multiple);

            MessageBox.Show("Возникла ошибка при инициализации компонентов .NET Framework 4.6.2", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }

        static void Search(DirectoryInfo dr, Regex file)
        {
            FileInfo[] fi = dr.GetFiles();
            foreach (FileInfo info in fi)
            {
                if (file.IsMatch(info.Name))
                {
                    Console.WriteLine("{0,-20} | {1}", info.Directory.Name, info.Name);
                }
            }
            DirectoryInfo[] dirs = dr.GetDirectories();
            foreach (DirectoryInfo directoryInfo in dirs)
            {
                Search(directoryInfo, file);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "GHbdtn";
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            MessageBox.Show("chekbox");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string str;
            FileStream fileStream = new FileStream(@"D:\Misha\recognization\Files\Emails.txt", FileMode.Open);
            StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8);
            //streamWriter.BaseStream.Seek(fileStream.Length, SeekOrigin.End);//запись в конец файла
            streamWriter.BaseStream.Seek(-5, SeekOrigin.End);//запись в конец файла
            str = textBox4 + "|" + textBox5;
            streamWriter.WriteLine(str);
            streamWriter.Close();
            fileStream.Close();
            textBox4.Text = "";
            textBox5.Text = "";
        }
    }
}
