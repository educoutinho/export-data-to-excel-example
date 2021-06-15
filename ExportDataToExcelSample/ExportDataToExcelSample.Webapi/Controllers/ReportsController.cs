using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace ExportDataToExcelSample.Webapi.Controllers
{
    [ApiController]
	[Route("[controller]")]
	public class ReportsController : ControllerBase
	{
		private readonly ILogger<ReportsController> _logger;

		public ReportsController(ILogger<ReportsController> logger)
		{
			_logger = logger;
		}

		[HttpGet]
		[Route("export")]
		[ProducesResponseType((int)HttpStatusCode.OK)]
		public async Task<FileResult> Export(string cookieName = null)
		{
			var byteArray = await Reports.CreateExcelFileAsync();

			return this.DownloadExcelFile(byteArray, "MyFile.xlsx", cookieName);
		}
	}

	public static class ControllerExtensions
	{
		public static FileResult DownloadExcelFile(this ControllerBase controller, byte[] byteArray, string fileName, string cookieName)
		{
			if (!string.IsNullOrEmpty(cookieName))
				controller.Response.Cookies.Append(cookieName, "true");

			return controller.File(byteArray, "application/Excel", fileName);
		}
	}

	public static class Reports
    {
		public static async Task<byte[]> CreateExcelFileAsync()
		{
			var workbook = new XSSFWorkbook();
			var sheet = workbook.CreateSheet("Plan 1");

			int rowNumber = 0;
			int colIndex;

			//---- HEADER

			var row = sheet.CreateRow(rowNumber);

			var styleHeader = workbook.CreateCellStyle();
			styleHeader.FillForegroundColor = HSSFColor.Grey25Percent.Index;
			styleHeader.FillPattern = FillPattern.SolidForeground;

			ICell cell;

            var columns = new List<ColumnInfo>
            {
                new ColumnInfo() { Name = "Nome", Width = 40 },
                new ColumnInfo() { Name = "Telefone", Width = 30 },
                new ColumnInfo() { Name = "Valor 1", Width = 10 },
                new ColumnInfo() { Name = "Valor 2", Width = 10 },
                new ColumnInfo() { Name = "Soma", Width = 10 }
            };

            for (int i = 0; i < columns.Count; i++)
			{
				cell = row.CreateCell(i);
				cell.SetCellValue(columns[i].Name);
				cell.CellStyle = styleHeader;
			}

			//---- row
			rowNumber++;
			colIndex = 0;
			row = sheet.CreateRow(rowNumber);
			row.CreateCell(colIndex++).SetCellValue("Eduardo");
			row.CreateCell(colIndex++).SetCellValue("111111");
			row.CreateCell(colIndex++).SetCellValue("10");
			row.CreateCell(colIndex++).SetCellValue("7");
			row.CreateCell(colIndex++).SetCellFormula("C2+D2");

			//---- row
			rowNumber++;
			colIndex = 0;
			row = sheet.CreateRow(rowNumber);
			row.CreateCell(colIndex++).SetCellValue("Coutinho");
			row.CreateCell(colIndex++).SetCellValue("222222");
			row.CreateCell(colIndex++).SetCellValue("1");
			row.CreateCell(colIndex++).SetCellValue("2");
			row.CreateCell(colIndex++).SetCellFormula("C3+D3");

			//Adiciona um número mínimo de linhas, para evitar o erro: O Excel encontrou conteúdo ilegível / Invalid or corrupt file (unreadable content)
			while (rowNumber < 20)
			{
				rowNumber++;
				row = sheet.CreateRow(rowNumber);
				row.CreateCell(0).SetCellValue(" ");
				row.CreateCell(1).SetCellValue(" ");
			}

			//Ajusta o tamanho das colunas
			for (int i = 0; i < columns.Count; i++)
				sheet.SetColumnWidth(i, columns[i].Width * 256);

			byte[] byteArray;
			using (var stream = new MemoryStream())
			{
				workbook.Write(stream);
				byteArray = stream.ToArray();
			}

			return await Task.FromResult(byteArray);
		}

		private class ColumnInfo
		{
			public string Name { get; set; }
			public int Width { get; set; }
		}
	}
}
