using HtmlAgilityPack;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;

namespace Detsad
{
    class Program
    {
        static void Main()
        {
            ParseSite parseSite = new ParseSite();
            Console.WriteLine("выполняется Хабаровск");
            parseSite.Khabarovsk();
            Console.WriteLine("выполняется сайт a2b2ru");
            parseSite.A2b2ru();
            Console.WriteLine("выполняется Севостопль");
            parseSite.Sevastopol();
            Console.WriteLine("выполняется сайт отечество");
            parseSite.Ote4estvo();
            Console.WriteLine("выполняется Ижевск");
            parseSite.Izhevsk();
            Console.WriteLine("выполняется Москва");
            parseSite.Moscow();
            parseSite.Emailres();
            parseSite.ExcelFile();
        }
    }
    class ParseSite
    {
        public static void DeleteRow(XSSFSheet sheetm, IRow rowm)
        {
            sheetm.RemoveRow(rowm);
            int rowIndex = rowm.RowNum;
            int lastRowNum = sheetm.LastRowNum;
            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheetm.ShiftRows(rowIndex + 1, lastRowNum, -1);
            }
        }
        List<Parse> Liste = new List<Parse>();
        public void Khabarovsk()
        {
            WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
            string s = client.DownloadString("https://edu.khabarovskadm.ru/obshchestvennoe-upravlenie/mun_preschool_institutions/index.php?ELEMENT_ID=84017");
            client.Dispose();
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(s);
            HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//tr");
            if (all != null)
            {
                foreach (HtmlNode n in all)
                {
                    HtmlDocument rows = new HtmlDocument();
                    rows.LoadHtml(n.InnerHtml);
                    HtmlNodeCollection cell = rows.DocumentNode.SelectNodes("//td");
                    if ((cell.Count > 1) && ((cell[2] != null) && (cell[6] != null)))
                    {
                        string name = cell[1].InnerText.Replace("&nbsp;", "").Replace("\r","").Replace("\n", "").Replace("    ","").Replace("  ", "").Replace("г.Хабаровска", "")
                            .Replace(" г. Хабаровска", "").Replace("образовательноеучреждение", "образовательное учреждение").Replace("садкомбинированного", "сад комбинированного")
                            .Replace("комбинированноговида", "комбинированноговида");
                        string email = cell[5].InnerText.Replace("&nbsp;", "").Replace("\r", "").Replace("\n", "").Replace("    ", "").Replace(" ","");
                        Liste.Add(new Parse { Name = name, Email = email });
                    }
                }
            }
            Liste.RemoveAt(11);
            }
        public void A2b2ru()
            {
                for (int i = 1; i <= 305; i++)
                {
                    WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
                    string s = client.DownloadString("http://a2b2.ru/kindergardens/page" + i + "/");
                    client.Dispose();
                    string name="", email="";
                    HtmlDocument doc = new HtmlDocument();
                    doc.LoadHtml(s);
                    HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//div[@class='list-items_item_description']");
                    if (all != null)
                    {
                        foreach (HtmlNode n in all)
                        {
                            HtmlDocument doccc = new HtmlDocument();
                            doccc.LoadHtml(n.InnerHtml);
                            HtmlNode html = doccc.DocumentNode.SelectSingleNode("//a");
                            if (html != null)
                            {
                                name = html.InnerText.Replace("\n","").Replace("...","").Replace("\t","").Replace("&raquo;", "").Replace("&laquo;", "");
                            }
                            HtmlNode al = doccc.DocumentNode.SelectSingleNode("//p");
                            if (al != null)
                            {
                                HtmlDocument docc = new HtmlDocument();
                                docc.LoadHtml(al.InnerHtml);
                                HtmlNode a = docc.DocumentNode.SelectSingleNode("//a");
                                if(a != null)
                                {
                                    email = a.InnerText;
                                }
                            }
                            if (((name != null) && (email != null)) && ((name !="") && (email != "")))
                            {
                                Liste.Add(new Parse { Name = name, Email = email });
                            }
                        }
                    }
                }
            }
        public void Sevastopol()
        {
            WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
            string s = client.DownloadString("https://eduface.ru/sites/list/region/1");
            HtmlDocument doc = new HtmlDocument();
            string name = "", email = "";
            doc.LoadHtml(s);
            HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//div[@class='accordion-style4-wraplink']");
            if (all != null)
            {
                foreach (HtmlNode n in all)
                {
                    HtmlDocument docc = new HtmlDocument();
                    docc.LoadHtml(n.InnerHtml);
                    if ((n.InnerText.IndexOf("Детский") != -1) && (n.InnerText.IndexOf("лагерь") == -1) && (n.InnerText.IndexOf("школа") == -1))
                    {
                        name = n.InnerText.Replace("\n","").Replace("\t","").Replace("\r","");
                        while (name[name.Length-1] == ' ')
                        {
                            name = name.Remove(name.Length - 1, 1);
                        }
                        while (name[0] == ' ')
                        {
                            name = name.Remove(0, 1);
                        }
                        HtmlNode a = docc.DocumentNode.SelectSingleNode("//a");
                        {
                            if (a.Attributes["href"] != null)
                            {
                                docc.LoadHtml(client.DownloadString(a.Attributes["href"].Value));
                                HtmlNode node = docc.DocumentNode.SelectSingleNode("//div[@class='sitename']");
                                if (node != null)
                                {
                                    docc.LoadHtml(node.InnerHtml);
                                    HtmlNode b = docc.DocumentNode.SelectSingleNode("//a");
                                    if (a.Attributes["href"] != null)
                                    {
                                        docc.LoadHtml(client.DownloadString(a.Attributes["href"].Value +"/home"));
                                    }
                                }
                                HtmlNodeCollection nodes = docc.DocumentNode.SelectNodes("//div");
                                if (nodes != null)
                                {
                                    foreach (HtmlNode p in nodes)
                                    {
                                        if ((p.InnerText.IndexOf('@') != -1) && (p.InnerText.IndexOf(' ') == -1))
                                        {
                                              email = p.InnerText.Replace("\n", "").Replace("\r", "").Replace("\t", "");
                                        }
                                    }
                                }
                            }
                        }
                        Liste.Add(new Parse { Name = name, Email = email });
                    }
                }
            }
        }
        public void Ote4estvo()
        {
            for (int i = 1; i <= 50; i++)
            {
                WebClient client = new WebClient() { Encoding = Encoding.GetEncoding(1251) };
                string s = client.DownloadString("http://www.ote4estvo.ru/sadik/russia/page/" + i + "/");
                client.Dispose();
                string name = "", email = "";
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(s);
                HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//div[@class='short-story-block']");
                if (all != null)
                {
                    foreach (HtmlNode n in all)
                    {
                        HtmlDocument docc = new HtmlDocument();
                        docc.LoadHtml(n.InnerHtml);
                        HtmlNode nodes = docc.DocumentNode.SelectSingleNode("//h4");
                        name = nodes.InnerText;
                        HtmlNode node = docc.DocumentNode.SelectSingleNode("//a");
                        if (node.Attributes["href"] != null)
                        {
                            docc.LoadHtml(client.DownloadString(node.Attributes["href"].Value));
                            HtmlNodeCollection nodees = docc.DocumentNode.SelectNodes("//div[@class='item']");
                            if (nodees != null) 
                            {
                                foreach (HtmlNode m in nodees)
                                {
                                    if(m.InnerText.IndexOf('@') != -1)
                                    {
                                        email = m.InnerText.Replace(" ","").Replace("Почта","").Replace("-","");
                                        Liste.Add(new Parse { Name = name, Email = email });
                                    }
                                }
                            }
                        }
                    }      
                }
            }
        }
        public void Izhevsk()
        {
            string uri = "http://www.izh.ru";
            WebClient client = new WebClient() { Encoding = Encoding.UTF8 };
            string s = client.DownloadString("http://www.izh.ru/i/info/15271.html");
            client.Dispose();
            string name = "", email = "";
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(s);
            HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//div[@class='chldlinkitem']");
            if (all != null)
            {
                foreach (HtmlNode n in all)
                {
                    HtmlDocument docc = new HtmlDocument();
                    docc.LoadHtml(n.InnerHtml);
                    HtmlNode node = docc.DocumentNode.SelectSingleNode("//a");
                    if (node.Attributes["href"] != null)
                    {
                        docc.LoadHtml(client.DownloadString(uri +node.Attributes["href"].Value));
                        HtmlNode nodees = docc.DocumentNode.SelectSingleNode("//div[@class='dep_wigets']");
                        if (nodees != null)
                        {
                            HtmlNode nodeess = docc.DocumentNode.SelectSingleNode("//span[@class='dep_name']");
                            if (nodeess != null)
                            {
                                name = nodeess.InnerText;
                            }
                            email = nodees.InnerText;
                            int first = email.IndexOf("Email: ");
                            int lengh;
                            if (email.IndexOf(".ru") != -1)
                            {
                                lengh = email.IndexOf(".ru") - first + 3;
                            }
                            else if (email.IndexOf(".net") != -1)
                            {
                                lengh = email.IndexOf(".net") - first + 4;
                            }
                            else
                            {
                               lengh = email.IndexOf(".com") - first + 4;
                            }
                            email = email.Substring(first, lengh).Replace("Email: ", "");
                            Liste.Add(new Parse { Name = name, Email = email });
                        }
                    }
                }
            }
        }
        public void Moscow()
        {
            XSSFWorkbook xssfwb;
            using (FileStream file = new FileStream("detsadmoscow.xlsx", FileMode.Open, FileAccess.Read))
            {
                xssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = xssfwb.GetSheetAt(0);

            for (int row = 0; row <= sheet.LastRowNum; row++)
            {
                var currentRow = sheet.GetRow(row);
                if (currentRow != null)
                {
                    string name = currentRow.GetCell(0).StringCellValue;
                    string email = currentRow.GetCell(1).StringCellValue;
                    if((name != "") && (email != ""))
                    {
                        Liste.Add(new Parse { Name = name, Email = email });
                    }
                }
            }
        }
        public void ExcelFile()
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            XSSFSheet sh = (XSSFSheet)wb.CreateSheet("Лист 1");
            int countColumn = 2;
            for (int i = 0; i < Liste.Count; i++)
            {
                var currentRow = sh.CreateRow(i);
                for (int j = 0; j < countColumn; j++)
                {
                    var currentCell = currentRow.CreateCell(j);
                    if (j == 0)
                    {
                        currentCell.SetCellValue(Liste[i].Name);
                    }
                    if (j == 1)
                    {
                        currentCell.SetCellValue(Liste[i].Email);

                    }
                }
            }

            for (int i = 0; i < sh.LastRowNum; i++)
            {
                for (int j = i+1; j < sh.LastRowNum; j++)
                {
                    var currentRow = sh.GetRow(i);
                    var currentRows = sh.GetRow(j);
                    if ((currentRow != null) && (currentRows != null))
                    {
                        if(currentRows.GetCell(1).StringCellValue == currentRow.GetCell(1).StringCellValue)
                        {
                            DeleteRow(sh, sh.GetRow(j));
                        }
                    }
                }
            }
            if (!File.Exists("d:\\Дет.сад.xlsx"))
            {
                File.Delete("d:\\Дет.сад.xlsx");
            }
            using (var fs = new FileStream("d:\\Дет.сад.xlsx", FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }
            Process.Start("d:\\Дет.сад.xlsx");
            Liste.Clear();
        }
        public void Emailres()
        {
            for (int i = 0; i<= Liste.Count; i++)
            {
                string email = Liste[i].Email;
                int first = email.IndexOf("Email: ");
                int lengh;
                if (email.IndexOf(".ru") != -1)
                {
                    lengh = email.IndexOf(".ru") - first + 3;
                }
                else if (email.IndexOf(".net") != -1)
                {
                    lengh = email.IndexOf(".net") - first + 4;
                }
                else if (email.IndexOf(".li") != -1)
                {
                    lengh = email.IndexOf(".li") - first + 3;
                }
                else
                {
                    lengh = email.IndexOf(".com") - first + 4;
                }
                email = email.Substring(first, lengh);
            }
        }
    }
    public class Parse
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }
}
