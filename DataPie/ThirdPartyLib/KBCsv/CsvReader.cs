namespace Kent.Boogaart.KBCsv
{
    using Kent.Boogaart.HelperTrinity;
    using Kent.Boogaart.HelperTrinity.Extensions;
    using Kent.Boogaart.KBCsv.Internal;
    using System;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Provides a means of reading CSV data.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The <c>CsvReader</c> class allows CSV data to be read from an underlying data source. The data source may be a <see cref="Stream"/>, <see cref="TextReader"/>, file, or just a 
    /// <see cref="String"/> containing CSV data. There are various constructors allowing for file-, <c>Stream</c>- and <c>TextWriter</c>- based data sources, whilst
    /// <see cref="FromCsvString"/> facilitates reading CSV data from a <see cref="String"/>.
    /// </para>
    /// <para>
    /// By default, CSV values are assumed to be separated with a comma (<c>,</c>) and optionally delimited with double quotes (<c>"</c>). If necessary, the <see cref="ValueSeparator"/>
    /// and <see cref="ValueDelimiter"/> properties enable these default to be customized.
    /// </para>
    /// <para>
    /// Leading and trailing white space is discarded by default. <see cref="PreserveLeadingWhiteSpace"/> and <see cref="PreserveTrailingWhiteSpace"/> can be used to override these
    /// defaults.
    /// </para>
    /// <para>
    /// A header record can be read, either by calling <see cref="ReadHeaderRecord"/> (or <see cref="ReadHeaderRecordAsync"/>) or by assigning one directly to the <see cref="HeaderRecord"/>
    /// property. The latter approach is particularly useful when the underlying data does not actually contain the header record, but the structure of the data is known. Any attempts
    /// to read or assign a header record after the first record has been read by a <c>CsvReader</c> will result in an exception.
    /// </para>
    /// <para>
    /// Data records can be read via <see cref="ReadDataRecord"/> or <see cref="ReadDataRecords"/> (or their async counterparts, <see cref="ReadDataRecordAsync"/> and
    /// <see cref="ReadDataRecordsAsync"/>). These methods provide instances of <see cref="DataRecord"/>, which enables you to get at the underlying CSV values. Values can be
    /// accessed by index, or - if a <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> has been assigned - by column names.
    /// </para>
    /// <para>
    /// Data can be skipped via <see cref="O:Kent.Boogaart.KBCsv.CsvReader.SkipRecord"/> or <see cref="O:Kent.Boogaart.KBCsv.CsvReader.SkipRecords"/> (or their async counterparts,
    /// <see cref="O:Kent.Boogaart.KBCsv.CsvReader.SkipRecordAsync"/> and <see cref="O:Kent.Boogaart.KBCsv.CsvReader.SkipRecordsAsync"/>). Skipping data is faster than reading (and discarding)
    /// it. Thus, if it is known that a number of records are not needed, it is recommended they be skipped rather than read.
    /// </para>
    /// <para>
    /// The <see cref="RecordNumber"/> property gives the total number of records read, including any header record.
    /// </para>
    /// </remarks>
    /// <threadsafety>
    /// A <c>CsvReader</c> cannot be used safely from multiple threads without synchronization. When using the <c>Async</c> methods, you should ensure that any prior task has
    /// completed before instigating another.
    /// </threadsafety>
    /// <example>
    /// <para>
    /// The following example uses <see cref="FromCsvString"/> to parse data in a <see cref="String"/>. Values in data records are accessed via an index:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadCSVFromString" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadCSVFromString" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example uses <see cref="FromCsvString"/> to parse data in a <see cref="String"/>. White space is preserved:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadCSVFromStringPreservingWhiteSpace" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadCSVFromStringPreservingWhiteSpace" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example reads data from a file, where values are tab-separated and delimited with single quotes:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadTabDelimitedDataFromFile" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadTabDelimitedDataFromFile" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example reads data from an underlying UTF-8 encoded <see cref="Stream"/>. It treats the first record as a <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/>, thus enabling
    /// the data to be accessed via column names.
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadCSVFromStream" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadCSVFromStream" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example reads data from a file. It explicitly assigns a <see cref="HeaderRecord"/> and skips every other record:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadCSVFromFileWithExplicitHeader" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadCSVFromFileWithExplicitHeader" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example reads data asynchronously from an underlying <see cref="TextReader"/>. It uses <see cref="ReadDataRecordsAsync"/> instead of
    /// <see cref="ReadDataRecordAsync"/> to avoid creating too fine-grained tasks. It also requests that the <see cref="TextReader"/> be left open when disposing
    /// the <c>CsvReader</c>. All data records are dumped to the console:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadCSVFromFileAsynchronously" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadCSVFromFileAsynchronously" lang="vb"/>
    /// </example>
    public partial class CsvReader : IDisposable
    {
        private static readonly ExceptionHelper exceptionHelper = new ExceptionHelper(typeof(CsvReader));
        private readonly CsvParser parser;
        private readonly bool leaveOpen;
        private readonly DataRecord[] buffer;
        private HeaderRecord headerRecord;
        private long recordNumber;
        private bool disposed;

        /// <summary>
        /// Initializes a new instance of the CsvReader class.
        /// </summary>
        /// <remarks>
        /// <paramref name="stream"/> is assumed to be encoded with <see cref="Encoding.UTF8"/>, and will be disposed when this <c>CsvReader</c> is disposed.
        /// </remarks>
        /// <param name="stream">
        /// A stream from which CSV data will be read.
        /// </param>
        public CsvReader(Stream stream)
            : this(stream, Constants.DefaultEncoding)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvReader class.
        /// </summary>
        /// <remarks>
        /// <paramref name="stream"/> will be disposed when this <c>CsvReader</c> is disposed.
        /// </remarks>
        /// <param name="stream">
        /// A stream from which CSV data will be read.
        /// </param>
        /// <param name="encoding">
        /// The encoding for CSV data within the stream.
        /// </param>
        public CsvReader(Stream stream, Encoding encoding)
            : this(new StreamReader(stream, encoding), false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvReader class.
        /// </summary>
        /// <param name="stream">
        /// A stream from which CSV data will be read.
        /// </param>
        /// <param name="encoding">
        /// The encoding of CSV data within the stream.
        /// </param>
        /// <param name="leaveOpen">
        /// If <see langword="true"/>, <paramref name="stream"/> will not be disposed when this <c>CsvReader</c> is disposed.
        /// </param>
        public CsvReader(Stream stream, Encoding encoding, bool leaveOpen)
            : this(new StreamReader(stream, encoding), leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvReader class.
        /// </summary>
        /// <remarks>
        /// <paramref name="textReader"/> will be disposed when this <c>CsvReader</c> is disposed.
        /// </remarks>
        /// <param name="textReader">
        /// The source of the CSV data.
        /// </param>
        public CsvReader(TextReader textReader)
            : this(textReader, false)
        {
            textReader.AssertNotNull("textReader");
        }

        /// <summary>
        /// Initializes a new instance of the CsvReader class.
        /// </summary>
        /// <param name="textReader">
        /// The source of the CSV data.
        /// </param>
        /// <param name="leaveOpen">
        /// If <see langword="true"/>, <paramref name="textReader"/> will not be disposed when this <c>CsvReader</c> is disposed.
        /// </param>
        public CsvReader(TextReader textReader, bool leaveOpen)
        {
            textReader.AssertNotNull("textReader");

            this.parser = new CsvParser(textReader);
            this.leaveOpen = leaveOpen;

            // used to parse singular records
            this.buffer = new DataRecord[1];
        }

        /// <summary>
        /// Gets or sets a value indicating whether leading white space should be preserved.
        /// </summary>
        public bool PreserveLeadingWhiteSpace
        {
            get
            {
                this.EnsureNotDisposed();
                return this.parser.PreserveLeadingWhiteSpace;
            }

            set
            {
                this.EnsureNotDisposed();
                this.parser.PreserveLeadingWhiteSpace = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether trailing white space should be preserved.
        /// </summary>
        public bool PreserveTrailingWhiteSpace
        {
            get
            {
                this.EnsureNotDisposed();
                return this.parser.PreserveTrailingWhiteSpace;
            }

            set
            {
                this.EnsureNotDisposed();
                this.parser.PreserveTrailingWhiteSpace = value;
            }
        }

        /// <summary>
        /// Gets or sets the character used to separate values within the CSV.
        /// </summary>
        /// <remarks>
        /// This property specifies what character is used to separate values within the CSV. The default value separator is a comma (<c>,</c>).
        /// </remarks>
        public char ValueSeparator
        {
            get
            {
                this.EnsureNotDisposed();
                return this.parser.ValueSeparator;
            }

            set
            {
                this.EnsureNotDisposed();
                this.parser.ValueSeparator = value;
            }
        }

        /// <summary>
        /// Gets or sets the character used to delimit values within the CSV.
        /// </summary>
        /// <remarks>
        /// <para>
        /// This property specifies what character is used to delimit values within the CSV. The default value delimiter is a double quote (<c>"</c>). If set to <see langword="null"/>, no character will
        /// be treated as a delimiter.
        /// </para>
        /// <para>
        /// Note that values within the CSV are not required to be delimited. However, delimiting a value allows it to contain the delimiter character itself, along with new line characters (ie. a multi-line value).
        /// </para>
        /// </remarks>
        public char? ValueDelimiter
        {
            get
            {
                this.EnsureNotDisposed();
                return this.parser.ValueDelimiter;
            }

            set
            {
                this.EnsureNotDisposed();
                this.parser.ValueDelimiter = value;
            }
        }

        /// <summary>
        /// Gets or sets the header record.
        /// </summary>
        /// <remarks>
        /// <para>
        /// This property will be automatically set if <see cref="ReadHeaderRecord"/> or <see cref="ReadHeaderRecordAsync"/> is called. Alternatively, it can be set explicitly in the case where the underlying
        /// data source does not itself contain a header record, but the structure of the data is known.
        /// </para>
        /// <para>
        /// Any <see cref="DataRecord"/> produced by this <c>CsvReader</c> will have its <see cref="DataRecord.HeaderRecord"/> property set to the value of this property. Setting a header record allows 
        /// the values within data records to be accessed both by index and by column name.
        /// <see cref="DataRecord"/>
        /// </para>
        /// <para>
        /// Any attempt to set this property when the first record has already been read will result in an exception.
        /// </para>
        /// </remarks>
        public HeaderRecord HeaderRecord
        {
            get
            {
                this.EnsureNotDisposed();
                return this.headerRecord;
            }

            set
            {
                this.EnsureNotDisposed();
                this.EnsureNotPassedFirstRecord();
                value.AssertNotNull("value");
                this.headerRecord = value;
            }
        }

        /// <summary>
        /// Gets the current record number.
        /// </summary>
        /// <remarks>
        /// This property gives the number of records that have been read by this <c>CsvReader</c>. This includes the header record, unless it is provided explicitly via the
        /// <see cref="HeaderRecord"/> property.
        /// </remarks>
        public long RecordNumber
        {
            get
            {
                this.EnsureNotDisposed();
                return this.recordNumber;
            }
        }

        /// <summary>
        /// Gets a value indicating whether there are more records to read.
        /// </summary>
        public bool HasMoreRecords
        {
            get
            {
                this.EnsureNotDisposed();
                return this.parser.HasMoreRecords;
            }
        }

        /// <summary>
        /// Creates a <c>CsvReader</c> that will read the CSV data in <paramref name="csv"/>.
        /// </summary>
        /// <param name="csv">
        /// The CSV data to read.
        /// </param>
        /// <returns>
        /// A <c>CsvReader</c> that will read the CSV in <paramref name="csv"/>.
        /// </returns>
        public static CsvReader FromCsvString(string csv)
        {
            csv.AssertNotNull("csv");
            return new CsvReader(new StringReader(csv), true);
        }

        /// <summary>
        /// Attempts to skip a record in the data, and increments <see cref="RecordNumber"/> if successful.
        /// </summary>
        /// <returns>
        /// <see langword="true"/> if the record is successfully skipped, or <see langword="false"/> if there are no more records to skip.
        /// </returns>
        public bool SkipRecord()
        {
            return this.SkipRecords(1, true) == 1;
        }

        /// <summary>
        /// Attempts to skip a record in the data, optionally incrementing <see cref="RecordNumber"/> if successful.
        /// </summary>
        /// <param name="incrementRecordNumber">
        /// <see langword="true"/> to increment <see cref="RecordNumber"/> upon successfully skipping a record.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the record is successfully skipped, or <see langword="false"/> if there are no more records to skip.
        /// </returns>
        public bool SkipRecord(bool incrementRecordNumber)
        {
            return this.SkipRecords(1, incrementRecordNumber) == 1;
        }

        /// <summary>
        /// Attempts to skip <paramref name="count"/> records in the data, and increments <see cref="RecordNumber"/> by the number of records actually skipped.
        /// </summary>
        /// <remarks>
        /// If there are fewer than <paramref name="count"/> records remaining in the data, this method will skip the remaining records and return the number of records actually skipped.
        /// </remarks>
        /// <param name="count">
        /// The number of records to skip.
        /// </param>
        /// <returns>
        /// The actual number of records skipped.
        /// </returns>
        public int SkipRecords(int count)
        {
            return this.SkipRecords(count, true);
        }

        /// <summary>
        /// Attempts to skip <paramref name="count"/> records in the data, optionally incrementing <see cref="RecordNumber"/> by the number of records actually skipped.
        /// </summary>
        /// <remarks>
        /// If there are fewer than <paramref name="count"/> records remaining in the data, this method will skip the remaining records and return the number of records actually skipped.
        /// </remarks>
        /// <param name="count">
        /// The number of records to skip.
        /// </param>
        /// <param name="incrementRecordNumber">
        /// <see langword="true"/> to increment <see cref="RecordNumber"/> by the number of records skipped.
        /// </param>
        /// <returns>
        /// The actual number of records skipped.
        /// </returns>
        public int SkipRecords(int count, bool incrementRecordNumber)
        {
            this.EnsureNotDisposed();
            var skipped = this.parser.SkipRecords(count);

            if (incrementRecordNumber)
            {
                this.recordNumber += skipped;
            }

            return skipped;
        }

        /// <summary>
        /// Reads the first record from the underlying CSV data and assigns it to <see cref="HeaderRecord"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If successful, all <see cref="DataRecord"/>s read by this <c>CsvReader</c> will have their <see cref="DataRecord.HeaderRecord"/> set accordingly.
        /// </para>
        /// <para>
        /// Any attempt to call this method when this <c>CsvReader</c> has already read a record will result in an exception.
        /// </para>
        /// </remarks>
        /// <returns>
        /// The <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> that was read, also available via the <see cref="HeaderRecord"/> property. If no records are left, this method returns <see langword="null"/>.
        /// </returns>
        public HeaderRecord ReadHeaderRecord()
        {
            this.EnsureNotDisposed();
            this.EnsureNotPassedFirstRecord();

            if (this.parser.ParseRecords(null, this.buffer, 0, 1) == 1)
            {
                ++this.recordNumber;
                this.headerRecord = new HeaderRecord(this.buffer[0]);
                return this.headerRecord;
            }

            return null;
        }

        /// <summary>
        /// Reads a <see cref="DataRecord"/> from the underlying CSV.
        /// </summary>
        /// <returns>
        /// The <see cref="DataRecord"/> that was read, or <see langword="null"/> if there are no more records to read.
        /// </returns>
        public DataRecord ReadDataRecord()
        {
            this.EnsureNotDisposed();

            if (this.parser.ParseRecords(this.headerRecord, this.buffer, 0, 1) == 1)
            {
                ++this.recordNumber;
                return this.buffer[0];
            }

            return null;
        }

        /// <summary>
        /// Reads at most <paramref name="count"/> <see cref="DataRecord"/>s and populates <paramref name="buffer"/> with them, beginning at index <paramref name="offset"/>.
        /// </summary>
        /// <remarks>
        /// When reading a lot of data, it is possible that better performance can be achieved by using this method.
        /// </remarks>
        /// <param name="buffer">
        /// The buffer to populate with the <see cref="DataRecord"/>s that are read.
        /// </param>
        /// <param name="offset">
        /// The offset into <paramref name="buffer"/> at which to start placing data records.
        /// </param>
        /// <param name="count">
        /// The maximum number of data records to read.
        /// </param>
        /// <returns>
        /// The number of data records actually read and stored in <paramref name="buffer"/>.
        /// </returns>
        public int ReadDataRecords(DataRecord[] buffer, int offset, int count)
        {
            this.EnsureNotDisposed();

            var read = this.parser.ParseRecords(this.headerRecord, buffer, offset, count);
            this.recordNumber += read;
            return read;
        }

        /// <summary>
        /// Closes this <c>CsvReader</c>.
        /// </summary>
        /// <remarks>
        /// This method is an alternative means of disposing the <c>CsvReader</c>. Generally one should prefer a <c>using</c> block to automatically dispose of the <c>CsvReader</c>.
        /// </remarks>
        public void Close()
        {
            this.Dispose();
        }

        /// <summary>
        /// Disposes of this <c>CsvReader</c>.
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
            this.disposed = true;
        }

        /// <summary>
        /// Disposes of this <c>CsvReader</c>.
        /// </summary>
        /// <remarks>
        /// Subclasses can override this method to supplement dispose logic.
        /// </remarks>
        /// <param name="disposing">
        /// <see langword="true"/> if being called in response to a <see cref="Dispose"/> call, otherwise <see langword="false"/>.
        /// </param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing && !this.leaveOpen)
            {
                this.parser.TextReader.Dispose();
            }
        }

        private void EnsureNotPassedFirstRecord()
        {
            if (this.RecordNumber > 0 || this.headerRecord != null)
            {
                throw exceptionHelper.Resolve("passedFirstRecord");
            }
        }

        private void EnsureNotDisposed()
        {
            if (this.disposed)
            {
                throw exceptionHelper.Resolve("disposed");
            }
        }
    }
}