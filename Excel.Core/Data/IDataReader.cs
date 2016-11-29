﻿// Type: System.Data.IDataReader
// Assembly: System.Data, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089
// Assembly location: C:\Windows\Microsoft.NET\Framework\v4.0.30319\System.Data.dll

using System;
using System.Threading.Tasks;

// ReSharper disable CheckNamespace
namespace ExcelDataReader.Portable.Data
// ReSharper restore CheckNamespace
{
    public interface IDataReader : IDisposable, IDataRecord
    {
        int Depth { get; }

        bool IsClosed { get; }

        int RecordsAffected { get; }

        void Close();

        //DataTable GetSchemaTable();

        bool NextResult();

		Task<bool> Read();
    }
}
