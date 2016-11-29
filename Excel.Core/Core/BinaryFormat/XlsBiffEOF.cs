#if LEGACY
using Excel;
#endif
namespace ExcelDataReader.Portable.Core.BinaryFormat
{
	/// <summary>
	/// Represents BIFF EOF resord
	/// </summary>
	internal class XlsBiffEOF : XlsBiffRecord
	{
		internal XlsBiffEOF(byte[] bytes, uint offset, ExcelBinaryReader reader)
			: base(bytes, offset, reader)
		{
		}
	}
}
