using HtmlAgilityPack;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Detsad
{
    class Program
    {
       
        static void Main()
        {
            ParseSite parseSite = new ParseSite();
            parseSite.Khabarovsk();
            parseSite.A2b2ru();
            parseSite.ExcelFile();
        }
    }
    class ParseSite
    {
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
                        string email = cell[5].InnerText.Replace("&nbsp;", "").Replace("\r", "").Replace("\n", "").Replace("    ", "");
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
                string s = client.DownloadString("http://a2b2.ru/kindergardens/page" + i);
                client.Dispose();
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(s);
                HtmlNodeCollection all = doc.DocumentNode.SelectNodes("//div[@class='list-items_item_description']");
                if (all != null)
                {
                    foreach (HtmlNode n in all)
                    {
                        HtmlDocument doccc = new HtmlDocument();
                        doccc.LoadHtml(n.InnerHtml);
                        HtmlNode al = doccc.DocumentNode.SelectSingleNode("//p");
                        if (al != null)
                        {
                            HtmlDocument docc = new HtmlDocument();
                            docc.LoadHtml(al.InnerHtml);
                            HtmlNode a = docc.DocumentNode.SelectSingleNode("//a");
                            if(a != null)
                            {
                                string str = a.InnerText;
                            }
                        }
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
                    sh.AutoSizeColumn(j);
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

    }
    public class Parse
    {
        public string Name { get; set; }
        public string Email { get; set; }
    }
}
