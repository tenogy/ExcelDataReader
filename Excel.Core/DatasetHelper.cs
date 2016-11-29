using Excel.Core;
using ExcelDataReader.Portable.Data;

namespace ExcelDataReader.Desktop.Portable
{
	public class DatasetHelper : IDatasetHelper
	{
		private Worksheet _currentWorksheet;

		public bool IsValid { get; set; }

		public Workbook Workbook { get; private set; }

		public void CreateNew()
		{
			Workbook = new Workbook();
		}

		public void CreateNewTable(string name)
		{
			_currentWorksheet = new Worksheet(name);
		}

		public void EndLoadTable()
		{
			Workbook.Sheets.Add(_currentWorksheet);
		}

		public void AddColumn(string columnName)
		{
			_currentWorksheet.AddColumn(columnName);
		}

		public void BeginLoadData()
		{
		}

		public void AddRow(params object[] values)
		{
			_currentWorksheet.AddRow(values);
		}

		public void DatasetLoadComplete()
		{
		}

		public void AddExtendedPropertyToTable(string propertyName, string propertyValue)
		{
			_currentWorksheet.ExtendedProperties.Add(propertyName, propertyValue);
		}
	}
}
