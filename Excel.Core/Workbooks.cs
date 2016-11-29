using System.Collections.Generic;

namespace Excel.Core
{
	public class WorksheetColumn
	{
		public string Name { get; private set; }

		public WorksheetColumn(string name)
		{
			Name = name;
		}
	}

	public class WorksheetRow
	{
		public object[] Values { get; private set; }

		public WorksheetRow(object[] values)
		{
			Values = values;
		}
	}

	public class Worksheet
	{
		public string Name { get; private set; }
		public IDictionary<string, string> ExtendedProperties { get; private set; }
		public ICollection<WorksheetRow> Rows { get; private set; }
		public ICollection<WorksheetColumn> Columns { get; private set; }

		public Worksheet(string name)
		{
			Name = name;
			ExtendedProperties = new Dictionary<string, string>();
			Rows = new HashSet<WorksheetRow>();
			Columns = new HashSet<WorksheetColumn>();
		}

		public void AddRow(params object[] values)
		{
			Rows.Add(new WorksheetRow(values));
		}

		public void AddColumn(string name)
		{
			Columns.Add(new WorksheetColumn(name));
		}
	}

	public class Workbook
	{
		public ICollection<Worksheet> Sheets { get; private set; }

		public Workbook()
		{
			Sheets = new HashSet<Worksheet>();
		}
	}
}
