using System;
using System.IO;
using System.Threading.Tasks;
using Excel;

namespace CoreTest
{
	public class Program
	{
		public static void Main(string[] args)
		{
			Run().Wait();
		}


		static async Task Run()
		{
			try
			{
				using (var stream = new FileStream(@"..\Excel.Tests.4.5\Resources\Test_BigFormatted.xlsx", FileMode.Open))
				{
					IExcelDataReader reader = await ExcelReaderFactory.CreateOpenXmlReader(stream);

					reader.IsFirstRowAsColumnNames = true;
					var ds = await reader.ReadAll();
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
	}
}
