using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Collections;
using System.Linq;
using ExcelDataReader.Portable.IO;
using ExcelDataReader.Portable.Log;

namespace ExcelDataReader.Portable.Core
{
	public class ZipWorker : IDisposable
	{
		private readonly IFileHelper fileHelper;

		#region Members and Properties

		private bool disposed;
		private bool isCleaned;

		private const string TMP = "TMP_Z";
		private const string FOLDER_xl = "xl";
		private const string FOLDER_worksheets = "worksheets";
		private const string FILE_sharedStrings = "sharedStrings.{0}";
		private const string FILE_styles = "styles.{0}";
		private const string FILE_workbook = "workbook.{0}";
		private const string FILE_sheet = "sheet{0}.{1}";
		private const string FOLDER_rels = "_rels";
		private const string FILE_rels = "workbook.{0}.rels";

		private string tempPath;
		private string exceptionMessage;
		private string xlPath;
		private string format = "xml";

		private bool isValid;
		private string folderName;
		private DirectoryInfo rootFolder;

		//private bool _isBinary12Format;

		/// <summary>
		/// Gets a value indicating whether this instance is valid.
		/// </summary>
		/// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
		public bool IsValid
		{
			get { return isValid; }
		}

		/// <summary>
		/// Gets the temp path for extracted files.
		/// </summary>
		/// <value>The temp path for extracted files.</value>
		public string TempPath
		{
			get { return tempPath; }
		}

		/// <summary>
		/// Gets the exception message.
		/// </summary>
		/// <value>The exception message.</value>
		public string ExceptionMessage
		{
			get { return exceptionMessage; }
		}

		#endregion

		public ZipWorker(IFileHelper fileHelper)
		{
			this.fileHelper = fileHelper;
		}

		/// <summary>
		/// Extracts the specified zip file stream.
		/// </summary>
		/// <param name="fileStream">The zip file stream.</param>
		/// <returns></returns>
		public async Task<bool> Extract(Stream fileStream)
		{
			if (null == fileStream) return false;

			CleanFromTemp(false);

			NewTempPath();

			isValid = true;

			ZipArchive zipFile = null;

			try
			{
				zipFile = new ZipArchive(fileStream);

				IEnumerator enumerator = zipFile.Entries.GetEnumerator();

				while (enumerator.MoveNext())
				{
					var entry = (ZipArchiveEntry)enumerator.Current;

					await ExtractZipEntry(zipFile, entry);
				}
			}
			catch (InvalidDataException ex)
			{
				isValid = false;
				exceptionMessage = ex.Message;

				CleanFromTemp(true);
			}
			catch (Exception ex)
			{
				CleanFromTemp(true); //true tells CleanFromTemp not to raise an IO Exception if this operation fails. If it did then the real error here would be masked
				throw;
			}
			finally
			{
				fileStream.Dispose();

				if (null != zipFile) zipFile.Dispose();
			}

			return isValid && CheckFolderTree();
		}

		/// <summary>
		/// Gets the shared strings stream.
		/// </summary>
		/// <returns></returns>
		public async Task<Stream> GetSharedStringsStream()
		{
			return await GetStream(Path.Combine(xlPath, string.Format(FILE_sharedStrings, format)));
		}

		/// <summary>
		/// Gets the styles stream.
		/// </summary>
		/// <returns></returns>
		public async Task<Stream> GetStylesStream()
		{
			return await GetStream(Path.Combine(xlPath, string.Format(FILE_styles, format)));
		}

		/// <summary>
		/// Gets the workbook stream.
		/// </summary>
		/// <returns></returns>
		public async Task<Stream> GetWorkbookStream()
		{
			return await GetStream(Path.Combine(xlPath, string.Format(FILE_workbook, format)));
		}

		/// <summary>
		/// Gets the worksheet stream.
		/// </summary>
		/// <param name="sheetId">The sheet id.</param>
		/// <returns></returns>
		public async Task<Stream> GetWorksheetStream(int sheetId)
		{
			return await GetStream(Path.Combine(
				Path.Combine(xlPath, FOLDER_worksheets),
				string.Format(FILE_sheet, sheetId, format)));
		}

		public async Task<Stream> GetWorksheetStream(string sheetPath)
		{
			//its possible sheetPath starts with /xl. in this case trim the /xl
			if (sheetPath.StartsWith("/xl/"))
				sheetPath = sheetPath.Substring(4);
			return await GetStream(Path.Combine(xlPath, sheetPath));
		}


		/// <summary>
		/// Gets the workbook rels stream.
		/// </summary>
		/// <returns></returns>
		public async Task<Stream> GetWorkbookRelsStream()
		{
			return await GetStream(Path.Combine(xlPath, Path.Combine(FOLDER_rels, string.Format(FILE_rels, format))));
		}

		private void CleanFromTemp(bool catchIoError)
		{
			if (string.IsNullOrEmpty(tempPath)) return;

			isCleaned = true;

			try
			{
				var exists = File.Exists(tempPath);
				if (exists)
				{
					Directory.Delete(Path.GetDirectoryName(tempPath), true);
				}
			}
			catch (IOException ex)
			{
				this.Log().Error(ex.Message);
				if (!catchIoError)
					throw;
			}

		}

		
		private async Task ExtractZipEntry(ZipArchive zipFile, ZipArchiveEntry entry)
		{
			if (string.IsNullOrEmpty(entry.Name)) return;

			//create the file
			var filePath = Path.Combine(tempPath, entry.FullName);
			string dir = Path.GetDirectoryName(filePath);
			if (!Directory.Exists(dir))
			{
				Directory.CreateDirectory(dir);
			}
			var file = File.Exists(filePath) ? File.OpenWrite(filePath) : File.Create(filePath);

			using (var stream = file)
			{
				using (var inputStream = entry.Open())
				{
					await inputStream.CopyToAsync(stream);
					await stream.FlushAsync();
				}
			}
		}

		private void NewTempPath()
		{
			var tempID = Guid.NewGuid().ToString("N");
			folderName = TMP + DateTime.Now.ToFileTimeUtc().ToString() + tempID;
			tempPath = Path.Combine(fileHelper.GetTempPath(), folderName);

			//ensure root folder created
			var rootExists = Directory.Exists(tempPath);
			if (!rootExists)
			{
				Directory.CreateDirectory(tempPath);
			}
			rootFolder = new DirectoryInfo(tempPath);
			isCleaned = false;

			this.Log().Debug("Using temp path {0}", tempPath);

		}

		private bool CheckFolderTree()
		{
			xlPath = Path.Combine(tempPath, FOLDER_xl);

			var existsXlPath = Directory.Exists(xlPath);
			var existsWorksheetPath = Directory.Exists(Path.Combine(xlPath, FOLDER_worksheets));
			var existsWorkbook = File.Exists(Path.Combine(xlPath, FILE_workbook));
			var existsStyles = File.Exists(Path.Combine(xlPath, FILE_styles));

			return existsXlPath &&
				existsWorksheetPath &&
				existsWorkbook &&
				existsStyles;
		}

		private async Task<Stream> GetStream(string filePath)
		{

			if (File.Exists(filePath))
			{
				return await Task.FromResult(File.OpenRead(filePath));
			}
			else
			{
				return null;
			}
		}

		#region IDisposable Members

		public void Dispose()
		{
			Dispose(true);

			GC.SuppressFinalize(this);
		}

		private void Dispose(bool disposing)
		{
			// Check to see if Dispose has already been called.
			if (!this.disposed)
			{
				if (disposing)
				{
					if (!isCleaned)
						CleanFromTemp(false);
				}

				disposed = true;
			}
		}

		~ZipWorker()
		{
			Dispose(false);
		}

		#endregion
	}
}