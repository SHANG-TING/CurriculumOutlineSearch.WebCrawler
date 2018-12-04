using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using Newtonsoft.Json;
using CurriculumOutlineSearch.Models;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;

namespace CurriculumOutlineSearch.WebCrawler
{
    class Program
    {
        static readonly string rootPath = Path.Combine("E:\\data", DateTime.Now.Ticks.ToString());

        static void Main(string[] args)
        {
            WebClient wc = new WebClient();
            wc.Proxy = null;

            List<AcademicNounCategoryView> categoryViews = GetCategoryViews();

            foreach (AcademicNounCategoryView categoryView in categoryViews)
            {
                string zipPath = Path.Combine(rootPath, $"{categoryView.Id}.zip");
                string extractPath = Path.Combine(rootPath, $"{categoryView.Id}");
                string jsonPath = Path.Combine(rootPath, $"{categoryView.Id}.json");

                wc.DownloadFile(categoryView.DownloadUrl, zipPath);

                ZipFile.ExtractToDirectory(zipPath, extractPath);
                {
                    JArray jArray = new JArray();

                    foreach (var fileName in Directory.GetFiles(extractPath, "*.xls", SearchOption.AllDirectories))
                    {
                        string filePath = Path.Combine(extractPath, fileName);

                        #region xls 轉檔為 xlsx
                        {
                            var app = new Microsoft.Office.Interop.Excel.Application();
                            var wb = app.Workbooks.Open(filePath);
                            wb.SaveAs(Filename: filePath = filePath + "x", FileFormat: Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                            wb.Close();
                            app.Quit();
                        }
                        #endregion

                        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                        using (ExcelPackage ep = new ExcelPackage(fs))
                        {
                            ExcelWorksheet sheet = ep.Workbook.Worksheets[1];
                            int startRowNumber = sheet.Dimension.Start.Row + 1;
                            int endRowNumber = sheet.Dimension.End.Row;
                            int startColumn = sheet.Dimension.Start.Column;
                            int endColumn = sheet.Dimension.End.Column;

                            Dictionary<int, string> propertyNameDictionary = Enumerable.Range(startColumn, endColumn)
                                .ToDictionary(x => x, x => sheet.Cells[1, x].Text);

                            foreach(int row in Enumerable.Range(startRowNumber, endRowNumber)) {
                                JObject jObject = new JObject();

                                foreach(int col in Enumerable.Range(startColumn, endColumn))
                                {
                                    string propertyName = propertyNameDictionary[col];
                                    string propertyValue = sheet.Cells[row, col].Text;

                                    JToken token = jObject[propertyName];

                                    if (token == null)
                                    {
                                        jObject.Add(propertyName, propertyValue);
                                    }
                                }

                                jArray.Add(jObject);
                            }
                        }
                    }

                    File.WriteAllText(jsonPath, JsonConvert.SerializeObject(jArray));
                }

                File.Delete(zipPath);
            }
        }

        static List<AcademicNounCategoryView> GetCategoryViews()
        {
            string url = "http://terms.naer.edu.tw";
            string path = Path.Combine(rootPath, "category.json");

            HtmlWeb web = new HtmlWeb();
            HtmlDocument doc = web.Load($"{url}/download/");

            List<AcademicNounCategoryView> categoryViews = new List<AcademicNounCategoryView>();

            foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//div[@class='list-tab-content']/ul/li/a"))
            {
                var href = node.GetAttributeValue("href", null);

                if (string.IsNullOrEmpty(href))
                {
                    continue;
                }

                var categoryView = new AcademicNounCategoryView
                {
                    Id = Convert.ToInt32(href.TrimEnd('/').Split('/').Last()),
                    Name = node.InnerText.Trim(),
                };

                categoryView.DownloadUrl = $"{url}{href}Term_{categoryView.Id}.zip/";

                categoryViews.Add(categoryView);
            }

            if (!Directory.Exists(rootPath))
            {
                Directory.CreateDirectory(rootPath);
            }

            if (!File.Exists(path))
            {
                File.WriteAllText(path, JsonConvert.SerializeObject(categoryViews));
            }

            return categoryViews;
        }
    }
}
