namespace Kent.Boogaart.KBCsv
{
    using Kent.Boogaart.HelperTrinity;
    using Kent.Boogaart.HelperTrinity.Extensions;
    using Kent.Boogaart.KBCsv.Internal;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;
    using System.Text;

    /// <summary>
    /// Provides a means of writing CSV data.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The <c>CsvWriter</c> class allows CSV data to be written to an underlying data sink. The data sink may be a <see cref="Stream"/>, <see cref="TextWriter"/>, or a file.
    /// There are various constructors allowing for file-, <c>Stream</c>- and <c>TextWriter</c>- based data sinks.
    /// </para>
    /// <para>
    /// By default, CSV values will be separated with a comma (<c>,</c>) and delimited where necessary with double quotes (<c>"</c>). If desired, the <see cref="ValueSeparator"/>
    /// and <see cref="ValueDelimiter"/> properties enable these defaults to be customized. In addition, delimiters around values can be included even when not necessary by setting
    /// the <see cref="ForceDelimit"/> property to <see langword="true"/>.
    /// </para>
    /// <para>
    /// Records can be written with <see cref="O:Kent.Boogaart.KBCsv.CsvWriter.WriteRecord"/> and <see cref="WriteRecords"/> methods (or their async counterparts,
    /// <see cref="O:Kent.Boogaart.KBCsv.CsvWriter.WriteRecordAsync"/> and <see cref="WriteRecordsAsync"/>). Doing so will increase <see cref="RecordNumber"/> accordingly.
    /// </para>
    /// </remarks>
    /// <threadsafety>
    /// A <c>CsvReader</c> cannot be used safely from multiple threads without synchronization. When using the <c>Async</c> methods, you should ensure that any prior task has
    /// completed before instigating another.
    /// </threadsafety>
    /// <example>
    /// <para>
    /// The following example writes CSV data to a <see cref="StringWriter"/>:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="WriteCSVToString" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="WriteCSVToString" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example writes CSV data (with delimiters forced) to a <see cref="MemoryStream"/> with ASCII encoding. The <see cref="MemoryStream"/> is left open when
    /// the <c>CsvWriter</c> is closed:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="WriteCSVToStreamWithForcedDelimiting" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="WriteCSVToStreamWithForcedDelimiting" lang="vb"/>
    /// </example>
    /// <example>
    /// <para>
    /// The following example asynchronously reads CSV from a file and asynchronously writes it to another. The data is written as tab-delimited with a single quote delimiter.
    /// A buffer is used to read and write data in blocks of records rather than one record at a time:
    /// </para>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.CSharp\Program.cs" region="ReadCSVFromFileAndWriteToTabDelimitedFile" lang="cs"/>
    /// <code source="..\Src\Kent.Boogaart.KBCsv.Examples.VB\Program.vb" region="ReadCSVFromFileAndWriteToTabDelimitedFile" lang="vb"/>
    /// </example>
    public partial class CsvWriter : IDisposable
    {
        private static readonly ExceptionHelper exceptionHelper = new ExceptionHelper(typeof(CsvWriter));
        private readonly TextWriter textWriter;
        private readonly bool leaveOpen;
        private readonly StringBuilder valueBuilder;
        private readonly StringBuilder bufferBuilder;
        private bool forceDelimit;
        private bool disposed;
        private char? valueDelimiter;
        private char valueSeparator;
        private long recordNumber;

        /// <summary>
        /// Initializes a new instance of the CsvWriter class.
        /// </summary>
        /// <remarks>
        /// <paramref name="stream"/> will be encoded with <see cref="System.Text.Encoding.UTF8"/>, and will be disposed when this <c>CsvWriter</c> is disposed.
        /// </remarks>
        /// <param name="stream">
        /// A stream to which CSV data will be written.
        /// </param>
        public CsvWriter(Stream stream)
            : this(stream, Constants.DefaultEncoding)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvWriter class.
        /// </summary>
        /// <remarks>
        /// <paramref name="stream"/> will be disposed when this <c>CsvWriter</c> is disposed.
        /// </remarks>
        /// <param name="stream">
        /// A stream to which CSV data will be written.
        /// </param>
        /// <param name="encoding">
        /// The encoding for CSV data written to <paramref name="stream"/>.
        /// </param>
        public CsvWriter(Stream stream, Encoding encoding)
            : this(stream, encoding, false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvWriter class.
        /// </summary>
        /// <remarks>
        /// <paramref name="stream"/> will be disposed when this <c>CsvWriter</c> is disposed.
        /// </remarks>
        /// <param name="stream">
        /// A stream to which CSV data will be written.
        /// </param>
        /// <param name="encoding">
        /// The encoding for CSV data written to <paramref name="stream"/>.
        /// </param>
        /// <param name="leaveOpen">
        /// If <see langword="true"/>, <paramref name="stream"/> will not be disposed when this <c>CsvWriter</c> is disposed.
        /// </param>
        public CsvWriter(Stream stream, Encoding encoding, bool leaveOpen)
            : this(new StreamWriter(stream, encoding), leaveOpen)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvWriter class.
        /// </summary>
        /// <remarks>
        /// <paramref name="textWriter"/> will be disposed when this <c>CsvWriter</c> is disposed.
        /// </remarks>
        /// <param name="textWriter">
        /// The target for the CSV data.
        /// </param>
        public CsvWriter(TextWriter textWriter)
            : this(textWriter, false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the CsvWriter class.
        /// </summary>
        /// <param name="textWriter">
        /// The target for the CSV data.
        /// </param>
        /// <param name="leaveOpen">
        /// If <see langword="true"/>, <paramref name="textWriter"/> will not be disposed when this <c>CsvWriter</c> is disposed.
        /// </param>
        public CsvWriter(TextWriter textWriter, bool leaveOpen)
        {
            textWriter.AssertNotNull("textWriter");

            this.textWriter = textWriter;
            this.leaveOpen = leaveOpen;
            this.valueBuilder = new StringBuilder(128);
            this.bufferBuilder = new StringBuilder(2048);
            this.valueSeparator = Constants.DefaultValueSeparator;
            this.valueDelimiter = Constants.DefaultValueDelimiter;
        }

        /// <summary>
        /// Gets the <see cref="System.Text.Encoding"/> being used when writing CSV data.
        /// </summary>
        public Encoding Encoding
        {
            get
            {
                this.EnsureNotDisposed();
                return this.textWriter.Encoding;
            }
        }

        /// <summary>
        /// Gets the current record number.
        /// </summary>
        /// <remarks>
        /// This property gives the number of records that have been written by this <c>CsvWriter</c>.
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
        /// Gets or sets a value indicating whether delimiters should always be written.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If this property is set to <see langword="true"/>, values will always be wrapped in <see cref="ValueDelimiter"/>s, even if they're not actually required. If <see langword="false"/>,
        /// values will only be delimited where necessary. That is, when they contain characters that require delimiting.
        /// </para>
        /// <para>
        /// Note that you cannot set this property to <see langword="true"/> if <see cref="ValueDelimiter"/> has been set to <see langword="null"/>.
        /// </para>
        /// </remarks>
        public bool ForceDelimit
        {
            get
            {
                this.EnsureNotDisposed();
                return this.forceDelimit;
            }

            set
            {
                this.EnsureNotDisposed();
                exceptionHelper.ResolveAndThrowIf(!this.valueDelimiter.HasValue && value, "valueDelimiterRequiredIfForceDelimitIsTrue");
                this.forceDelimit = value;
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
                return this.valueSeparator;
            }

            set
            {
                this.EnsureNotDisposed();
                exceptionHelper.ResolveAndThrowIf(value == this.valueDelimiter, "valueSeparatorAndDelimiterCannotMatch");
                this.valueSeparator = value;
            }
        }

        /// <summary>
        /// Gets or sets the character used to delimit values within the CSV.
        /// </summary>
        /// <remarks>
        /// <para>
        /// This property specifies what character is used to delimit values within the CSV. The default value delimiter is a double quote (<c>"</c>).
        /// If set to <see langword="null"/>, values will never be delimited. Should they require delimiting in order to be valid CSV, an exception will
        /// be thrown instead.
        /// </para>
        /// <para>
        /// If <see cref="ForceDelimit"/> is <see langword="false"/>, delimiters will only be written where necessary. Note that <see cref="ForceDelimit"/>
        /// cannot be <see langword="true"/> if <see cref="ValueDelimiter"/> is <see langword="null"/>.
        /// </para>
        /// </remarks>
        public char? ValueDelimiter
        {
            get
            {
                this.EnsureNotDisposed();
                return this.valueDelimiter;
            }

            set
            {
                this.EnsureNotDisposed();
                exceptionHelper.ResolveAndThrowIf(value == this.valueSeparator, "valueSeparatorAndDelimiterCannotMatch");
                exceptionHelper.ResolveAndThrowIf(!value.HasValue && this.forceDelimit, "valueDelimiterRequiredIfForceDelimitIsTrue");
                this.valueDelimiter = value;
            }
        }

        /// <summary>
        /// Gets or sets the <see cref="String"/> used to separate lines within the CSV.
        /// </summary>
        /// <remarks>
        /// By default, this property is set to <see cref="Environment.NewLine"/>.
        /// </remarks>
        public string NewLine
        {
            get
            {
                this.EnsureNotDisposed();
                return this.textWriter.NewLine;
            }

            set
            {
                this.EnsureNotDisposed();
                this.textWriter.NewLine = value;
            }
        }

        /// <summary>
        /// Writes a record to this <c>CsvWriter</c>.
        /// </summary>
        /// <remarks>
        /// All values within <paramref name="record"/> are written in the order they appear.
        /// </remarks>
        /// <param name="record">
        /// The record to write.
        /// </param>
        public void WriteRecord(RecordBase record)
        {
            Debug.Assert(this.bufferBuilder.Length == 0, "Expecting buffer to be empty.");

            this.EnsureNotDisposed();
            record.AssertNotNull("record");
            this.WriteRecordToBuffer(record);
            this.FlushBufferToTextWriter();
        }

        /// <summary>
        /// Writes a record to this <c>CsvWriter</c>.
        /// </summary>
        /// <remarks>
        /// Any <see langword="null"/> values will be written as empty strings.
        /// </remarks>
        /// <param name="values">
        /// The values comprising the record.
        /// </param>
        public void WriteRecord(params string[] values)
        {
            Debug.Assert(this.bufferBuilder.Length == 0, "Expecting buffer to be empty.");

            this.EnsureNotDisposed();
            values.AssertNotNull("values");
            this.WriteRecordToBuffer(values);
            this.FlushBufferToTextWriter();
        }

        /// <summary>
        /// Writes a record to this <c>CsvWriter</c>.
        /// </summary>
        /// <remarks>
        /// Any <see langword="null"/> values will be written as empty strings.
        /// </remarks>
        /// <param name="values">
        /// The values comprising the record.
        /// </param>
        public void WriteRecord(IEnumerable<string> values)
        {
            Debug.Assert(this.bufferBuilder.Length == 0, "Expecting buffer to be empty.");

            this.EnsureNotDisposed();
            values.AssertNotNull("values");
            this.WriteRecordToBuffer(values);
            this.FlushBufferToTextWriter();
        }

        /// <summary>
        /// Writes <paramref name="length"/> records to this <c>CsvWriter</c>.
        /// </summary>
        /// <remarks>
        /// When writing a lot of data, it is possible that better performance can be achieved by using this method.
        /// </remarks>
        /// <param name="buffer">
        /// The buffer containing the records to be written.
        /// </param>
        /// <param name="offset">
        /// The offset into <paramref name="buffer"/> from which the first record will be obtained.
        /// </param>
        /// <param name="length">
        /// The number of records to write.
        /// </param>
        public void WriteRecords(RecordBase[] buffer, int offset, int length)
        {
            Debug.Assert(this.bufferBuilder.Length == 0, "Expecting buffer to be empty.");

            this.EnsureNotDisposed();
            buffer.AssertNotNull("buffer");
            exceptionHelper.ResolveAndThrowIf(offset < 0 || offset >= buffer.Length, "invalidOffset");
            exceptionHelper.ResolveAndThrowIf(offset + length > buffer.Length, "invalidLength");

            for (var i = offset; i < offset + length; ++i)
            {
                var record = buffer[i];
                exceptionHelper.ResolveAndThrowIf(record == null, "recordNull");
                this.WriteRecordToBuffer(record);
            }

            // we only flush once, when all records have been written
            this.FlushBufferToTextWriter();
        }

        /// <summary>
        /// Flushes this <c>CsvWriter</c>.
        /// </summary>
        /// <remarks>
        /// This method can be used to flush the underlying <see cref="TextWriter"/> to which this <c>CsvWriter</c> is writing data.
        /// </remarks>
        public void Flush()
        {
            Debug.Assert(this.bufferBuilder.Length == 0, "Expecting buffer to be empty.");

            this.EnsureNotDisposed();
            this.textWriter.Flush();
        }

        /// <summary>
        /// Closes this <c>CsvWriter</c>.
        /// </summary>
        /// <remarks>
        /// This method is an alternative means of disposing the <c>CsvWriter</c>. Generally one should prefer a <c>using</c> block to automatically dispose of the <c>CsvWriter</c>.
        /// </remarks>
        public void Close()
        {
            this.Dispose();
        }

        /// <summary>
        /// Disposes of this <c>CsvWriter</c>.
        /// </summary>
        public void Dispose()
        {
            if (this.disposed)
            {
                return;
            }

            this.Dispose(true);
            GC.SuppressFinalize(this);
            this.disposed = true;
        }

        /// <summary>
        /// Disposes of this <c>CsvWriter</c>.
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
                this.textWriter.Dispose();
            }
        }

        /// <summary>
        /// Writes a character to the buffer being used to construct a value.
        /// </summary>
        /// <remarks>
        /// <para>
        /// By default, this method will ensure <paramref name="delimit"/> is <see langword="true"/> if <paramref name="ch"/> is <see cref="ValueSeparator"/>, <see cref="ValueDelimiter"/>, or an end of line character.
        /// If <paramref name="ch"/> is <see cref="ValueDelimiter"/>, it will repeat it so that it is escaped in the CSV value.
        /// </para>
        /// <para>
        /// Subclasses can override this behavior if necessary, but should take care to ensure that the resultant value in <paramref name="buffer"/> is valid for the scenario being addressed. <paramref name="delimit"/>
        /// must be set to <see langword="true"/> if the value needs to be delimited prior to writing it to the underlying <see cref="TextWriter"/>.
        /// </para>
        /// </remarks>
        /// <param name="buffer">
        /// The buffer containing the value being constructed.
        /// </param>
        /// <param name="ch">
        /// The character to append.
        /// </param>
        /// <param name="delimit">
        /// <see langword="true"/> if the value should be delimited.
        /// </param>
        protected virtual void WriteCharToBuffer(StringBuilder buffer, char ch, ref bool delimit)
        {
            delimit = delimit || ch == this.valueSeparator || ch == this.valueDelimiter || ch == Constants.CR || ch == Constants.LF;
            buffer.Append(ch);

            if (ch == this.valueDelimiter)
            {
                // value delimiter is repeated so it is escaped
                buffer.Append(ch);
            }
        }

        private static bool IsWhitespace(char ch)
        {
            return ch == Constants.Space || ch == Constants.Tab;
        }

        private void EnsureNotDisposed()
        {
            if (this.disposed)
            {
                throw exceptionHelper.Resolve("disposed");
            }
        }

        // synchronously push whatever's in the buffer to the text writer and reset the buffer
        private void FlushBufferToTextWriter()
        {
            this.textWriter.Write(this.bufferBuilder.ToString());
            this.bufferBuilder.Length = 0;
        }

        // write a record to the buffer builder
        private void WriteRecordToBuffer(IEnumerable<string> values)
        {
            Debug.Assert(values != null, "Expecting non-null values.");

            var firstValue = true;

            foreach (var value in values)
            {
                if (!firstValue)
                {
                    this.bufferBuilder.Append(this.valueSeparator);
                }

                this.WriteValueToBuffer(value);
                firstValue = false;
            }

            this.bufferBuilder.Append(this.NewLine);
            ++this.recordNumber;
        }

        // write value to the buffer, escaping embedded delimiters, and wrapping in delimiters as necessary
        private void WriteValueToBuffer(string value)
        {
            var delimit = this.forceDelimit;
            this.valueBuilder.Length = 0;

            if (!string.IsNullOrEmpty(value))
            {
                // delimit to preserve white-space at the beginning or end of the value
                if (IsWhitespace(value[0]) || IsWhitespace(value[value.Length - 1]))
                {
                    delimit = true;
                }

                for (var i = 0; i < value.Length; ++i)
                {
                    this.WriteCharToBuffer(this.valueBuilder, value[i], ref delimit);
                }
            }

            if (delimit)
            {
                exceptionHelper.ResolveAndThrowIf(!this.valueDelimiter.HasValue, "valueRequiresDelimiting", this.valueBuilder.ToString());
                this.bufferBuilder.Append(this.valueDelimiter).Append(this.valueBuilder).Append(this.valueDelimiter);
            }
            else
            {
                this.bufferBuilder.Append(this.valueBuilder);
            }
        }
    }
}