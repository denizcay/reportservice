using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Npgsql;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;

namespace ReportProject.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        public ReportController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet]
        public IActionResult Get()
        {
            byte[] fileContents;
            ExcelPackage.LicenseContext = LicenseContext.Commercial;


            using (var package = new ExcelPackage())
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.Columns.Width = 18;                

                #region Header Row
                workSheet.Cells[1, 1].Value = "Auto Sender ID";
                workSheet.Cells[1, 1].Style.Font.Size = 12;
                workSheet.Cells[1, 1].Style.Font.Bold = true;
                workSheet.Cells[1, 1].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[1, 2].Value = "# of Users";
                workSheet.Cells[1, 2].Style.Font.Size = 12;
                workSheet.Cells[1, 2].Style.Font.Bold = true;                
                workSheet.Cells[1, 2].Style.Border.Top.Style = ExcelBorderStyle.Hair;

                workSheet.Cells[1, 3].Value = "Items Replied";
                workSheet.Cells[1, 3].Style.Font.Size = 12;
                workSheet.Cells[1, 3].Style.Font.Bold = true;
                workSheet.Cells[1, 3].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                #endregion
                
            string query = "SELECT AutoSenderId, count(questionid), COUNT(answertext) FROM reportsource GROUP BY autosenderid ORDER BY autosenderid";
            DataTable table = new DataTable();
            string sqlDataSource = _configuration.GetConnectionString("AppCon");
            NpgsqlDataReader reader;
            using (NpgsqlConnection con = new NpgsqlConnection(sqlDataSource))
            {
                con.Open();
                using(NpgsqlCommand command = new NpgsqlCommand(query, con))
                {
                    reader = command.ExecuteReader();
                        table.Load(reader);

                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            for (int j = 0; j < 3; j++)
                            {
                                object[] temp = table.Rows[i].ItemArray;
                                workSheet.Cells[i + 2 , j + 1].Value = temp[j];
                            }
                        }

                        reader.Close();
                    con.Close();
                }
            }

                fileContents = package.GetAsByteArray();
            }

            if (fileContents == null || fileContents.Length == 0) return NotFound();

            return File(
                fileContents: fileContents,
                contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                fileDownloadName: "Report.xlsx"
                );
        }
    }
}
