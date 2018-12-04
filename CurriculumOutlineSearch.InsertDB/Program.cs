using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using CurriculumOutlineSearch.Models;
using Newtonsoft.Json;

namespace CurriculumOutlineSearch.InsertDB
{
    class Program
    {
        static readonly string RootPath = "E:\\data\\636794542983175298";
        static readonly string ConnectionString = "Database=CurriculumOutlineSearch;Server=.;Integrated Security=false;User ID=sa;Password=aa1234@@;";

        static void Main(string[] args)
        {
            UpdateAcadmicNounCategory();
            UpdateAcadmicNoun();
        }

        static void UpdateAcadmicNounCategory()
        {
            var dt = new DataTable();

            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("DownloadUrl", typeof(string));

            var filePath = Path.Combine(RootPath, "category.json");
            var text = File.ReadAllText(filePath);
            var json = JsonConvert.DeserializeObject<IEnumerable<AcademicNounCategoryView>>(text);
            
            foreach (var data in json)
            {
                var row = dt.NewRow();
                row["Id"] = data.Id;
                row["Name"] = data.Name;
                row["DownloadUrl"] = data.DownloadUrl;

                dt.Rows.Add(row);
            }

            using (var tx = new TransactionScope())
            {
                using (var sql = new SqlConnection(ConnectionString))
                {
                    sql.Open();

                    using (var cmd = new SqlCommand("Delete From dbo.AcademicNounCategory", sql))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    using (var sqlBulkCopy = new SqlBulkCopy(sql))
                    {
                        sqlBulkCopy.DestinationTableName = "dbo.AcademicNounCategory";
                        sqlBulkCopy.WriteToServer(dt);
                    }
                }

                tx.Complete();
            }
        }

        static void UpdateAcadmicNoun()
        {
            var dt = new DataTable();

            dt.Columns.Add("Id", typeof(int));
            dt.Columns.Add("SystemNumber", typeof(int));
            dt.Columns.Add("EnglishName", typeof(string));
            dt.Columns.Add("ChineseName", typeof(string));
            dt.Columns.Add("CategoryId", typeof(int));

            var allFilePath = Directory.GetFiles(RootPath, "*.json", SearchOption.AllDirectories)
                .Where(x => !x.EndsWith("category.json"))
                .ToArray();
            
            for (var i = 1; i < allFilePath.Length; i++ )
            {
                var filePath = allFilePath[i];
                var categoryId = Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath));
                var text = File.ReadAllText(filePath);
                var json = JsonConvert.DeserializeObject<IEnumerable<dynamic>>(text)
                    .Where(x => !string.IsNullOrWhiteSpace(Convert.ToString(x["系統編號"])))
                    .ToList();

                foreach (var item in json)
                {
                    var row = dt.NewRow();
                    row["Id"] = i;
                    row["SystemNumber"] = Convert.ToInt32(item["系統編號"]);
                    row["EnglishName"] = Convert.ToString(item["英文名稱"]);
                    row["ChineseName"] = Convert.ToString(item["中文名稱"]);
                    row["CategoryId"] = categoryId;

                    dt.Rows.Add(row);
                    i++;
                }
            }


            using (var tx = new TransactionScope())
            {
                using (var sql = new SqlConnection(ConnectionString))
                {
                    sql.Open();

                    using (var cmd = new SqlCommand("Delete From dbo.AcademicNoun", sql))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    using (var sqlBulkCopy = new SqlBulkCopy(sql))
                    {
                        sqlBulkCopy.DestinationTableName = "dbo.AcademicNoun";
                        sqlBulkCopy.WriteToServer(dt);
                    }
                }

                tx.Complete();
            }
        }
    }
}
