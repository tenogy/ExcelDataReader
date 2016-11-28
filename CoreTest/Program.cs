using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CoreTest
{
	public class Program
	{
		public static void Main(string[] args)
		{
			using (var stream = new FileStream(@"..\Excel.Tests.4.5\Resources\Test_BigFormatted.xlsx", FileMode.Open))
			{
				IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(stream);

				reader.IsFirstRowAsColumnNames = firstRowNamesCheckBox.Checked;
				var ds = reader.AsDataSet();
			}
		}
	}
}
