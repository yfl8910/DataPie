using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Kent.Boogaart.HelperTrinity;
using Kent.Boogaart.HelperTrinity.Extensions;

#if !SILVERLIGHT
using System.Data;
#endif

namespace Kent.Boogaart.KBCsv
{
    /// <summary>
    /// Provides a mechanism via which CSV data can be easily parsed.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The <c>CsvReader</c> class allows CSV data to be read and parsed from any stream-based source. By default, CSV values are assumed to be separated
    /// by commas (<c>,</c>) and delimited by double quotes (<c>"</c>). If necessary, custom characters can be specified when creating the
    /// <c>CsvReader</c>.
    /// </para>
    /// <para>
    /// The number of records that have been parsed so far is exposed via the <see cref="RecordNumber"/> property. Reading a header record does not affect
    /// this property.
    /// </para>
    /// <para>
    /// Once a <c>CsvReader</c> instance is created, a header record can optionally be parsed. If a header record exists, it must be parsed first thing
    /// via the <see cref="ReadHeaderRecord"/> method. Upon successfully parsing the header record, it is exposed via the <see cref="HeaderRecord"/>
    /// property. An instance of <see cref="HeaderRecord"/> represents the header record.
    /// </para>
    /// <para>
    /// If a header record does not exist in the CSV data but the details of the header are known, the <see cref="HeaderRecord"/> property can be used
    /// to explicitly supply header information. Only one of <see cref="HeaderRecord"/> or <see cref="ReadHeaderRecord"/> can be successfully used to set
    /// header information.
    /// </para>
    /// <para>
    /// Data records can be read using the overloaded <see cref="ReadDataRecord"/> method. Data records are represented as instances of
    /// <see cref="DataRecord"/>. Alternatively, <see cref="ReadDataRecordAsStrings"/> returns data as an array of type <see cref="string"/>.
    /// </para>
    /// <para>
    /// The header and data records can be skipped using either of the <see cref="SkipRecord"/> methods. By default, skipping a record increments the
    /// record number as exposed by the <see cref="RecordNumber"/> property. However, you can optionally specify <see langword="false"/> when invoking
    /// <see cref="SkipRecord(bool)"/> to ensure the record number is not incremented.
    /// </para>
    /// <para>
    /// The properties <see cref="DataRecords"/> and <see cref="DataRecordsAsStrings"/> provide a convenient way to obtain
    /// <see cref="IEnumerable&lt;T&gt;"/> instances that iterate over CSV data as instances of <see cref="DataRecord"/> and <see cref="string"/>[]
    /// respectively.
    /// </para>
    /// </remarks>
    /// <threadsafety>
    /// The <c>CsvReader</c> class does not lock internally. Therefore, it is unsafe to share an instance across threads without implementing your own
    /// synchronisation solution.
    /// </threadsafety>
    /// <example>
    /// <para>
    /// The following example uses one of the <see cref="FromCsvString"/> overloads to parse and screen dump the CSV data in a <c>string</c> instance:
    /// </para>
    /// <para>
    /// <code lang="C#">
    /// <![CDATA[
    /// string csvData = GetCsvData();
    /// 
    /// using (CsvReader reader = CsvReader.FromCsvString(csvData)) {
    ///		foreach (DataRecord record in reader.DataRecords) {
    ///			System.Console.WriteLine(record.ToString());
    ///		}
    /// }
    /// ]]>
    /// </code>
    /// </para>
    /// <para>
    /// <code lang="vb">
    /// <![CDATA[
    /// Dim csvData As String = GetCsvData
    /// Dim reader As CsvReader = Nothing
    /// 
    /// Try 
    ///		reader = CsvReader.FromCsvString(csvData)
    ///		
    ///		For Each record As DataRecord In reader.DataRecords
    ///			System.Console.WriteLine(record.ToString)
    ///		Next
    /// Finally
    ///		If (Not reader Is Nothing) Then
    ///			reader.Close
    ///		End If
    /// End Try
    /// ]]>
    /// </code>
    /// </para>
    /// </example>
    /// <example>
    /// <para>
    /// The following example parses CSV data in a file but specifies that the separator character is a colon (<c>:</c>), not a comma (<c>,</c>). In
    /// addition, it reads the first record as a header and references data via column name instead of index:
    /// </para>
    /// <para>
    /// <code lang="C#">
    /// <![CDATA[
    /// using (CsvReader reader = new CsvReader(@"C:\data.csv")) {
    ///		reader.ValueSeparator = ':';
    ///		reader.ReadHeaderRecord();
    ///		DataRecord record = null;
    ///		
    ///		while ((record = reader.ReadDataRecord()) != null) {
    ///			System.Console.WriteLine("{0} is {1} years old", record["Name"], record["Age"]);
    ///		}
    /// }
    /// ]]>
    /// </code>
    /// </para>
    /// <para>
    /// <code lang="vb">
    /// <![CDATA[
    /// Dim reader As CsvReader = Nothing
    /// 
    /// Try
    ///		reader = New CsvReader("C:\data.csv")
    ///		reader.ValueSeparator = ":"c
    ///		reader.ReadHeaderRecord()
    ///		Dim record As DataRecord = reader.ReadDataRecord
    /// 
    ///		While (Not record Is Nothing)
    ///			System.Console.WriteLine("{0} is {1} years old", record("Name"), record("Age"))
    ///			record = reader.ReadDataRecord
    ///		End While
    /// Finally
    ///		If (Not reader Is Nothing) Then
    ///			reader.Close()
    ///		End If
    /// End Try
    /// ]]>
    /// </code>
    /// </para>
    /// <para>
    /// The data in <c>C:\data.csv</c> might look like this, for example:
    /// <code>
    /// Name,Gender,Age
    /// Kent,M,25
    /// Belinda,F,26
    /// Tempany,F,0
    /// </code>
    /// </para>
    /// </example>
    /// <example>
    /// <para>
    /// The following example reads the data in a file and outputs every 10th record, along with a count of the records that have been printed:
    /// </para>
    /// <para>
    /// <code lang="C#">
    /// <![CDATA[
    /// using (CsvReader reader = new CsvReader(@"C:\data.csv")) {
    ///		DataRecord record = null;
    /// 
    ///		while (true) {
    ///			reader.SkipRecords(9, false);
    ///			record = reader.ReadDataRecord();
    /// 
    ///			if (record != null) {
    ///				Console.WriteLine("Record {0}: {1}", reader.RecordNumber, record);
    ///			} else {
    ///				break;
    ///			}
    ///		}
    /// }
    /// ]]>
    /// </code>
    /// </para>
    /// <para>
    /// <code lang="vb">
    /// <![CDATA[
    /// Dim reader As CsvReader = Nothing
    /// 
    /// Try
    ///		reader = New CsvReader("C:\data.csv")
    ///		Dim record As DataRecord = Nothing
    /// 
    ///		While True
    ///			reader.SkipRecords(9, False)
    ///			record = reader.ReadDataRecord
    /// 
    ///			If (Not record Is Nothing) Then
    ///				Console.WriteLine("Record {0}: {1}", reader.RecordNumber, record)
    ///				record = reader.ReadDataRecord
    ///			Else
    ///				Exit While
    ///			End If
    ///		End While
    /// Finally
    ///		If (Not reader Is Nothing) Then
    ///			reader.Close()
    ///		End If
    /// End Try
    /// ]]>
    /// </code>
    /// </para>
    /// </example>
    /// <example>
    /// <para>
    /// The following example fills a <c>DataSet</c> with CSV data. The table is called "csv-data". White space before each value is preserved:
    /// </para>
    /// <para>
    /// <code lang="C#">
    /// <![CDATA[
    /// using (CsvReader reader = new CsvReader(@"C:\data.csv")) {
    ///		reader.PreserveLeadingWhiteSpace = true;
    ///		reader.ReadHeaderRecord();
    ///		DataSet dataSet = new DataSet();
    ///		reader.Fill(dataSet, "csv-data");
    /// }
    /// ]]>
    /// </code>
    /// </para>
    /// <para>
    /// <code lang="vb">
    /// <![CDATA[
    /// Dim reader As CsvReader = Nothing
    /// 
    /// Try
    ///		reader = New CsvReader("C:\data.csv")
    ///		reader.PreserveLeadingWhiteSpace = True
    ///		reader.ReadHeaderRecord()
    ///		Dim dataSet As DataSet = New DataSet
    ///		reader.Fill(dataSet, "csv-data")
    /// Finally
    ///		If (Not reader Is Nothing) Then
    ///			reader.Close()
    ///		End If
    /// End Try
    /// ]]>
    /// </code>
    /// </para>
    /// </example>
    public class CsvReader : IDisposable
    {
        private static readonly ExceptionHelper _exceptionHelper = new ExceptionHelper(typeof(CsvReader));

        private static readonly Encoding defaultEncoding =
#if !SILVERLIGHT
 Encoding.Default;
#else
 Encoding.UTF8;
#endif

        /// <summary>
        /// The instance of <see cref="CsvParser"/> being used to parse the CSV data.
        /// </summary>
        private readonly CsvParser _parser;

        /// <summary>
        /// See <see cref="HeaderRecord"/>.
        /// </summary>
        private HeaderRecord _headerRecord;

        /// <summary>
        /// See <see cref="RecordNumber"/>.
        /// </summary>
        private long _recordNumber;

        /// <summary>
        /// Set to <see langword="true"/> when this object is disposed.
        /// </summary>
        private bool _disposed;

#if !SILVERLIGHT
        /// <summary>
        /// The default table name used during a <see cref="Fill"/> operation.
        /// </summary>
        public static readonly string DefaultTableName = "Table";
#endif

        /// <summary>
        /// Gets or sets a value indicating whether leading whitespace is to be preserved during parsing.
        /// </summary>
        public bool PreserveLeadingWhiteSpace
        {
            get
            {
                EnsureNotDisposed();
                return _parser.PreserveLeadingWhiteSpace;
            }
            set
            {
                EnsureNotDisposed();
                _parser.PreserveLeadingWhiteSpace = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether trailing whitespace is to be preserved during parsing.
        /// </summary>
        public bool PreserveTrailingWhiteSpace
        {
            get
            {
                EnsureNotDisposed();
                return _parser.PreserveTrailingWhiteSpace;
            }
            set
            {
                EnsureNotDisposed();
                _parser.PreserveTrailingWhiteSpace = value;
            }
        }

        /// <summary>
        /// Gets or sets the character placed between values in the CSV data.
        /// </summary>
        /// <remarks>
        /// This property can be used to determine what character the <c>CsvReader</c> will treat as a value separator whilst parsing CSV data. The
        /// default separator is a comma (<c>,</c>).
        /// </remarks>
        public char ValueSeparator
        {
            get
            {
                EnsureNotDisposed();
                return _parser.ValueSeparator;
            }
            set
            {
                EnsureNotDisposed();
                _parser.ValueSeparator = value;
            }
        }

        /// <summary>
        /// Gets the character possibly placed around values in the CSV data.
        /// </summary>
        /// <remarks>
        /// <para>
        /// This property can be used to determine what character the <c>CsvReader</c> will treat as a demarcation around values. The default delimiter
        /// is a double quote (<c>"</c>).
        /// </para>
        /// <para>
        /// Note that the CSV parser does not require values to be delimited with this character. However, values that are delimited can contain
        /// special characters such as <see cref="ValueSeparator"/> and <see cref="Environment.NewLine"/>.
        /// </para>
        /// </remarks>
        public char ValueDelimiter
        {
            get
            {
                EnsureNotDisposed();
                return _parser.ValueDelimiter;
            }
            set
            {
                EnsureNotDisposed();
                _parser.ValueDelimiter = value;
            }
        }

        /// <summary>
        /// Gets or sets the CSV header for this reader.
        /// </summary>
        /// <value>
        /// The CSV header record for this reader, or <see langword="null"/> if no header record applies.
        /// </value>
        /// <remarks>
        /// This property yields the CSV header for the current CSV data. It can also be used to explicitly set a header record rather than using
        /// <see cref="ReadHeaderRecord"/>.
        /// </remarks>
        public HeaderRecord HeaderRecord
        {
            get
            {
                EnsureNotDisposed();
                return _headerRecord;
            }
            set
            {
                EnsureNotDisposed();
                value.AssertNotNull("value");
                _exceptionHelper.ResolveAndThrowIf(_headerRecord != null, "HeaderRecord.set.header-record-already-set");
                _exceptionHelper.ResolveAndThrowIf(_parser.PassedFirstRecord, "HeaderRecord.set.first-record-already-read");

                _parser.PassedFirstRecord = true;
                _headerRecord = value;
            }
        }

        /// <summary>
        /// Gets the current record number.
        /// </summary>
        /// <remarks>
        /// This property gives the number of records that the <c>CsvReader</c> has parsed. The CSV header does not count. That is, calling
        /// <see cref="ReadHeaderRecord"/> will not increment this property. Only successful calls to <see cref="ReadDataRecord"/> (and related methods)
        /// will increment this property.
        /// </remarks>
        public long RecordNumber
        {
            get
            {
                EnsureNotDisposed();
                return _recordNumber;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this <c>CsvReader</c> has more CSV records to read in addition to those already read.
        /// </summary>
        /// <remarks>
        /// This property can be used to determine whether a call to <see cref="ReadDataRecord"/>, for example, will successfully retrieve a record. It
        /// does not affect the state of the <c>CsvReader</c>.
        /// </remarks>
        public bool HasMoreRecords
        {
            get
            {
                EnsureNotDisposed();
                return _parser.HasMoreRecords;
            }
        }

        /// <summary>
        /// Gets an instance of <see cref="IEnumerable&lt;T&gt;"/> that can be used to iterate over the CSV data records from the current position of
        /// this reader and its end.
        /// </summary>
        /// <remarks>
        /// This property can be used to iterate over the data records in this <c>CsvReader</c>.
        /// </remarks>
        public IEnumerable<DataRecord> DataRecords
        {
            get
            {
                EnsureNotDisposed();
                DataRecord record = null;

                while ((record = ReadDataRecord()) != null)
                {
                    yield return record;
                }
            }
        }

        /// <summary>
        /// Gets an instance of <see cref="IEnumerable&lt;T&gt;"/> that can be used to iterate over the CSV data in this reader as arrays of
        /// <c>string</c>s.
        /// </summary>
        /// <remarks>
        /// This property can be used to iterate over the data in this <c>CsvReader</c>.
        /// </remarks>
        public IEnumerable<string[]> DataRecordsAsStrings
        {
            get
            {
                EnsureNotDisposed();
                string[] record = null;

                while ((record = ReadDataRecordAsStrings()) != null)
                {
                    yield return record;
                }
            }
        }

        /// <summary>
        /// Creates an instance of <c>CsvReader</c> that uses <paramref name="csv"/> as its CSV data.
        /// </summary>
        /// <remarks>
        /// This method can be used to create a <c>CsvReader</c> that reads the CSV data in the supplied <c>string</c>.
        /// </remarks>
        /// <param name="csv">
        /// The CSV data to be read.
        /// </param>
        /// <returns>
        /// An instance of <c>CsvReader</c> that will read <paramref name="csv"/>
        /// </returns>
        public static CsvReader FromCsvString(string csv)
        {
            csv.AssertNotNull("csv");

            return new CsvReader(new StringReader(csv));
        }

        /// <summary>
        /// Constructs and initialises an instance of <c>CsvReader</c> based on the information provided.
        /// </summary>
        /// <param name="stream">
        /// The stream to read CSV data from.
        /// </param>
        public CsvReader(Stream stream)
            : this(stream, defaultEncoding)
        {
        }

        /// <summary>
        /// Constructs and initialises an instance of <c>CsvReader</c> based on the information provided.
        /// </summary>
        /// <param name="stream">
        /// The stream to read CSV data from.
        /// </param>
        /// <param name="encoding">
        /// The encoding for the data in <paramref name="stream"/>.
        /// </param>
        public CsvReader(Stream stream, Encoding encoding)
            : this(new StreamReader(stream, encoding))
        {
        }

        /// <summary>
        /// Constructs and initialises an instance of <c>CsvReader</c> based on the information provided.
        /// </summary>
        /// <param name="path">
        /// The full path to the file containing CSV data.
        /// </param>
        public CsvReader(string path)
            : this(path, defaultEncoding)
        {
        }

        /// <summary>
        /// Constructs and initialises an instance of <c>CsvReader</c> based on the information provided.
        /// </summary>
        /// <param name="path">
        /// The full path to the file containing CSV data.
        /// </param>
        /// <param name="encoding">
        /// The encoding for the data in the file at <paramref name="path"/>.
        /// </param>
        public CsvReader(string path, Encoding encoding)
            : this(new StreamReader(path, encoding))
        {
        }

        /// <summary>
        /// Constructs and initialises an instance of <c>CsvReader</c> based on the information provided.
        /// </summary>
        /// <param name="reader">
        /// The instance of <see cref="TextReader"/> from which CSV data can be read.
        /// </param>
        public CsvReader(TextReader reader)
        {
            _parser = new CsvParser(reader);
        }

        /// <summary>
        /// Disposes of this <c>CsvReader</c> instance.
        /// </summary>
        void IDisposable.Dispose()
        {
            Close();
            Dispose(true);
        }

        /// <summary>
        /// Allows sub classes to implement disposing logic.
        /// </summary>
        /// <param name="disposing">
        /// <see langword="true"/> if this method is being called in response to a <see cref="Dispose"/> call, or <see langword="false"/> if
        /// it is being called during finalization.
        /// </param>
        protected virtual void Dispose(bool disposing)
        {
        }

        /// <summary>
        /// Closes this <c>CsvReader</c> instance and releases all resources acquired by it.
        /// </summary>
        /// <remarks>
        /// Once an instance of <c>CsvReader</c> is no longer needed, call this method to immediately release any resources. Closing a <c>CsvReader</c> is equivalent to
        /// disposing of it via a C# <c>using</c> block.
        /// </remarks>
        public void Close()
        {
            if (_parser != null)
            {
                _parser.Close();
            }

            _disposed = true;
        }

        /// <summary>
        /// Reads the first record from the CSV source and treats it as the header for the remainder of the parse.
        /// </summary>
        /// <remarks>
        /// <para>
        /// CSV data sources can optionally include a header record, which contains the names of the columns in the data. This method can be called to
        /// ensure that the first record of CSV data is treated as a header record. After making a successful call to this method, the header record is
        /// available via the <see cref="HeaderRecord"/> property on this <c>CsvReader</c>, but also from the <see cref="DataRecord.HeaderRecord"/>
        /// property in each instance of <see cref="DataRecord"/> created by this <c>CsvReader</c>.
        /// </para>
        /// <para>
        /// If a header record is present and read via this method, data can be accessed via column name as well as column index. For example:
        /// <code>
        /// <![CDATA[
        /// using (CsvReader reader = new CsvReader(@"C:\data.csv")) {
        ///		reader.ReadHeaderRecord();
        ///		DataRecord dataRecord = reader.ReadDataRecord();
        ///		string age = dataRecord["Age"];
        /// }
        /// ]]>
        /// </code>
        /// </para>
        /// </remarks>
        /// <returns>
        /// The header for the CSV source, or <see langword="null"/> if no record was found to treat as a header.
        /// </returns>
        public HeaderRecord ReadHeaderRecord()
        {
            _exceptionHelper.ResolveAndThrowIf(_parser.PassedFirstRecord, "HeaderRecord.set.first-record-already-read");

            HeaderRecord headerRecord = HeaderRecord.FromParser(_parser);

            if (headerRecord != null)
            {
                //reset this flag since the first record read was simply the header
                _parser.PassedFirstRecord = false;
                HeaderRecord = headerRecord;
            }

            return HeaderRecord;
        }

        /// <summary>
        /// Reads the next data record and returns it. <see langword="null"/> is returned if no more records are found in the CSV source.
        /// </summary>
        /// <remarks>
        /// This method will parse the next record and return it as an instance of <see cref="DataRecord"/>. If the convenience of the <c>DataRecord</c> class is not needed,
        /// you can use the <see cref="ReadDataRecordAsStrings"/> to gain a slight performance benefit.
        /// </remarks>
        /// <returns>
        /// The next CSV data record, or <see langword="null"/> if no more records were found in the CSV source.
        /// </returns>
        public DataRecord ReadDataRecord()
        {
            EnsureNotDisposed();

            DataRecord retVal = DataRecord.FromParser(_headerRecord, _parser);

            if (retVal != null)
            {
                ++_recordNumber;
            }

            return retVal;
        }

        /// <summary>
        /// Reads the next data record and returns it as an array of strings.
        /// </summary>
        /// <remarks>
        /// This method may be useful if the convenience of a <see cref="DataRecord"/> instance is not required. Invoking this method instead of
        /// <see cref="ReadDataRecord"/> yields a slight performance benefit.
        /// </remarks>
        /// <returns>
        /// An array of type <c>string</c> containing the values in the next record, or <see langword="null"/> if no record was found.
        /// </returns>
        public string[] ReadDataRecordAsStrings()
        {
            EnsureNotDisposed();

            string[] retVal = _parser.ParseRecord();

            if (retVal != null)
            {
                ++_recordNumber;
            }

            return retVal;
        }

        /// <summary>
        /// Reads and returns all CSV data records as a collection of <c>DataRecord</c> instances.
        /// </summary>
        /// <remarks>
        /// This method reads the CSV data into memory. For more efficient reading, it is often better to use the <see cref="DataRecords"/> property, which yields an
        /// enumerator that only keeps one record in memory at a time. Alternatively, you can restrict the number of records read by using the
        /// <see cref="ReadDataRecords(int)"/> overload.
        /// </remarks>
        /// <returns>
        /// A collection of all CSV records as <c>DataRecord</c> instances. This method will never return <see langword="null"/>.
        /// </returns>
        public ICollection<DataRecord> ReadDataRecords()
        {
            EnsureNotDisposed();
            List<DataRecord> retVal = new List<DataRecord>();
            DataRecord record = null;

            while ((record = DataRecord.FromParser(_headerRecord, _parser)) != null)
            {
                retVal.Add(record);
            }

            return retVal;
        }

        /// <summary>
        /// Reads and returns all CSV data records as <c>DataRecord</c> instances up to a specified maximum number of records.
        /// </summary>
        /// <param name="maximumRecords">
        /// The maximum number of records to read.
        /// </param>
        /// <remarks>
        /// This method reads the CSV data into memory. For more efficient reading, it is often better to use the <see cref="DataRecords"/> property, which yields an
        /// enumerator that only keeps one record in memory at a time.
        /// </remarks>
        /// <returns>
        /// A collection of records whose length will never surpass <paramref name="maximumRecords"/>. If there are insufficient records in the CSV data,
        /// this method will retrieve the remaining records. This method will never return <see langword="null"/>.
        /// </returns>
        public ICollection<DataRecord> ReadDataRecords(int maximumRecords)
        {
            EnsureNotDisposed();
            _exceptionHelper.ResolveAndThrowIf(maximumRecords < 0, "ReadDataRecords.maximumRecords-less-than-zero");

            List<DataRecord> retVal = new List<DataRecord>();
            int numRead = 0;

            while (numRead < maximumRecords)
            {
                DataRecord record = DataRecord.FromParser(_headerRecord, _parser);

                if (record != null)
                {
                    ++_recordNumber;
                }
                else
                {
                    break;
                }

                ++numRead;
                retVal.Add(record);
            }

            return retVal;
        }

        /// <summary>
        /// Reads and returns all CSV data records as <c>string</c> arrays.
        /// </summary>
        /// <remarks>
        /// This method reads the CSV data into memory. For more efficient reading, it is often better to use the <see cref="DataRecordsAsStrings"/> property, which yields an
        /// enumerator that only keeps one record in memory at a time. Alternatively, you can restrict the number of records read by using the
        /// <see cref="ReadDataRecordsAsStrings(int)"/> overload.
        /// </remarks>
        /// <returns>
        /// A collection of all CSV records as <c>string</c> arrays. This method will never return <see langword="null"/>.
        /// </returns>
        public ICollection<string[]> ReadDataRecordsAsStrings()
        {
            EnsureNotDisposed();
            List<string[]> retVal = new List<string[]>();
            string[] record = null;

            while ((record = _parser.ParseRecord()) != null)
            {
                retVal.Add(record);
            }

            return retVal;
        }

        /// <summary>
        /// Reads and returns all CSV data records as <c>string</c> arrays up to a specified maximum number of records.
        /// </summary>
        /// <param name="maximumRecords">
        /// The maximum number of records to read.
        /// </param>
        /// <remarks>
        /// This method reads the CSV data into memory. For more efficient reading, it is often better to use the <see cref="DataRecordsAsStrings"/> property, which yields an
        /// enumerator that only keeps one record in memory at a time.
        /// </remarks>
        /// <returns>
        /// A collection of records whose length will never surpass <paramref name="maximumRecords"/>. If there are insufficient records in the CSV data,
        /// this method will retrieve the remaining records. This method will never return <see langword="null"/>.
        /// </returns>
        public ICollection<string[]> ReadDataRecordsAsStrings(int maximumRecords)
        {
            EnsureNotDisposed();
            _exceptionHelper.ResolveAndThrowIf(maximumRecords < 0, "ReadDataRecords.maximumRecords-less-than-zero");

            List<string[]> retVal = new List<string[]>();
            int numRead = 0;

            while (numRead < maximumRecords)
            {
                string[] record = _parser.ParseRecord();

                if (record != null)
                {
                    ++_recordNumber;
                }
                else
                {
                    break;
                }

                ++numRead;
                retVal.Add(record);
            }

            return retVal;
        }

        /// <summary>
        /// Skips the next CSV record in the data source.
        /// </summary>
        /// <remarks>
        /// This method skips over a CSV record and increments the record counter if the skip was successful. To avoid incrementing the record counter,
        /// use <see cref="SkipRecord(bool)"/>.
        /// </remarks>
        /// <returns>
        /// <see langword="true"/> if the record was skipped, or <see langword="false"/> if there is no record to skip.
        /// </returns>
        public bool SkipRecord()
        {
            EnsureNotDisposed();

            if (_parser.SkipRecord())
            {
                ++_recordNumber;
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Skips the next CSV record and optionally increments the record counter.
        /// </summary>
        /// <param name="incrementRecordNumber">
        /// If <see langword="true"/>, the <see cref="RecordNumber"/> property will be incremented on a successful skip operation.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the record was skipped, or <see langword="false"/> if there is no record to skip.
        /// </returns>
        public bool SkipRecord(bool incrementRecordNumber)
        {
            EnsureNotDisposed();

            if (_parser.SkipRecord())
            {
                if (incrementRecordNumber)
                {
                    ++_recordNumber;
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Skips the specified number of records and increments the record counter for each record skipped.
        /// </summary>
        /// <param name="number">
        /// The maximum number of records to skip.
        /// </param>
        /// <returns>
        /// The actual number of records skipped. If this is less than <paramref name="number"/> then the end of the CSV data was reached.
        /// </returns>
        public int SkipRecords(int number)
        {
            EnsureNotDisposed();
            _exceptionHelper.ResolveAndThrowIf(number < 0, "SkipRecords.number-less-than-zero");

            int retVal = 0;

            while (retVal < number)
            {
                if (_parser.SkipRecord())
                {
                    ++_recordNumber;
                    ++retVal;
                }
                else
                {
                    break;
                }
            }

            return retVal;
        }

        /// <summary>
        /// Skips the specified number of records, optionally incrementing the record counter for each record skipped.
        /// </summary>
        /// <param name="number">
        /// The maximum number of records to skip.
        /// </param>
        /// <param name="incrementRecordNumber">
        /// If <see langword="true"/>, the <see cref="RecordNumber"/> property will be incremented for each record skipped.
        /// </param>
        /// <returns>
        /// The actual number of records skipped. If this is less than <paramref name="number"/> then the end of the CSV data was reached.
        /// </returns>
        public int SkipRecords(int number, bool incrementRecordNumber)
        {
            EnsureNotDisposed();
            _exceptionHelper.ResolveAndThrowIf(number < 0, "SkipRecords.number-less-than-zero");

            int retVal = 0;

            while (retVal < number)
            {
                if (_parser.SkipRecord())
                {
                    if (incrementRecordNumber)
                    {
                        ++_recordNumber;
                    }

                    ++retVal;
                }
                else
                {
                    break;
                }
            }

            return retVal;
        }

#if !SILVERLIGHT

        /// <summary>
        /// Fills the specified <see cref="DataSet"/> with CSV data.
        /// </summary>
        /// <remarks>
        /// The header record for the <c>CsvReader</c> must be set prior to invoking this method. The data read must conform to the header record. That
        /// is, if a record is found with more columns than specified by the header record, an exception will be thrown.
        /// </remarks>
        /// <param name="dataSet">
        /// The <c>DataSet</c> to be filled.
        /// </param>
        /// <returns>
        /// The number of records read and stored in <paramref name="dataSet"/>.
        /// </returns>
        public int Fill(DataSet dataSet)
        {
            return Fill(dataSet, DefaultTableName, -1, false);
        }

        /// <summary>
        /// Fills the specified <see cref="DataSet"/> with CSV data.
        /// </summary>
        /// <remarks>
        /// The header record for the <c>CsvReader</c> must be set prior to invoking this method. The data read must conform to the header record. That
        /// is, if a record is found with more columns than specified by the header record, an exception will be thrown.
        /// </remarks>
        /// <param name="dataSet">
        /// The <c>DataSet</c> to be filled.
        /// </param>
        /// <param name="tableName">
        /// The name for the table created in the <c>DataSet</c> that holds the CSV data.
        /// </param>
        /// <returns>
        /// The number of records read and stored in <paramref name="dataSet"/>.
        /// </returns>
        public int Fill(DataSet dataSet, string tableName)
        {
            return Fill(dataSet, tableName, -1, false);
        }

        /// <summary>
        /// Fills the specified <see cref="DataSet"/> with CSV data.
        /// </summary>
        /// <remarks>
        /// The header record for the <c>CsvReader</c> must be set prior to invoking this method. The data read must conform to the header record. That
        /// is, if a record is found with more columns than specified by the header record, an exception will be thrown.
        /// </remarks>
        /// <param name="dataSet">
        /// The <c>DataSet</c> to be filled.
        /// </param>
        /// <param name="maximumRecords">
        /// The maximum number of records to read.
        /// </param>
        /// <returns>
        /// The number of records read and stored in <paramref name="dataSet"/>.
        /// </returns>
        public int Fill(DataSet dataSet, int maximumRecords)
        {
            return Fill(dataSet, DefaultTableName, maximumRecords, true);
        }

        /// <summary>
        /// Fills the specified <see cref="DataSet"/> with CSV data.
        /// </summary>
        /// <remarks>
        /// The header record for the <c>CsvReader</c> must be set prior to invoking this method. The data read must conform to the header record. That
        /// is, if a record is found with more columns than specified by the header record, an exception will be thrown.
        /// </remarks>
        /// <param name="dataSet">
        /// The <c>DataSet</c> to be filled.
        /// </param>
        /// <param name="tableName">
        /// The name for the table created in the <c>DataSet</c> that holds the CSV data.
        /// </param>
        /// <param name="maximumRecords">
        /// The maximum number of records to read.
        /// </param>
        /// <returns>
        /// The number of records read and stored in <paramref name="dataSet"/>.
        /// </returns>
        public int Fill(DataSet dataSet, string tableName, int maximumRecords)
        {
            return Fill(dataSet, tableName, maximumRecords, true);
        }

        /// <summary>
        /// Fills the specified <see cref="DataSet"/> with CSV data.
        /// </summary>
        /// <remarks>
        /// The header record for the <c>CsvReader</c> must be set prior to invoking this method. The data read must conform to the header record. That
        /// is, if a record is found with more columns than specified by the header record, an exception will be thrown.
        /// </remarks>
        /// <param name="dataSet">
        /// The <c>DataSet</c> to be filled.
        /// </param>
        /// <param name="tableName">
        /// The name for the table created in the <c>DataSet</c> that holds the CSV data.
        /// </param>
        /// <param name="maximumRecords">
        /// The maximum number of records to read. Only relevant if <paramref name="useMaximum"/> is <see langword="true"/>.
        /// </param>
        /// <param name="useMaximum">
        /// If <see langword="true"/>, <paramref name="maximumRecords"/> takes affect. Otherwise, no limit is imposed on the number of records read.
        /// </param>
        /// <returns>
        /// The number of records read and stored in <paramref name="dataSet"/>.
        /// </returns>
        private int Fill(DataSet dataSet, string tableName, int maximumRecords, bool useMaximum)
        {
            EnsureNotDisposed();
            dataSet.AssertNotNull("dataSet");
            tableName.AssertNotNull("tableName");
            _exceptionHelper.ResolveAndThrowIf(useMaximum && (maximumRecords < 0), "Fill.maximumRecords-less-than-zero", "maximumRecords");
            _exceptionHelper.ResolveAndThrowIf(_headerRecord == null, "Fill.no-header-record-set");

            DataTable table = dataSet.Tables.Add(tableName);

            //set up the table columns based on the header record
            foreach (string column in _headerRecord.Values)
            {
                table.Columns.Add(column);
            }

            int num = 0;

            while (!useMaximum || (num < maximumRecords))
            {
                string[] record = ReadDataRecordAsStrings();

                if (record != null)
                {
                    _exceptionHelper.ResolveAndThrowIf(record.Length > _headerRecord.Values.Count, "Fill.too-many-columns-in-record", record.Length, _headerRecord.Values.Count);

                    string[] recordAsStrings = new string[record.Length];
                    record.CopyTo(recordAsStrings, 0);
                    table.Rows.Add(recordAsStrings);
                    ++num;
                }
                else
                {
                    break;
                }
            }

            return num;
        }

#endif

        /// <summary>
        /// Makes sure the object isn't disposed and, if so, throws an exception.
        /// </summary>
        private void EnsureNotDisposed()
        {
            _exceptionHelper.ResolveAndThrowIf(_disposed, "disposed");
        }
    }
}
