namespace Kent.Boogaart.KBCsv
{
    using System.Security;
    using System.Threading.Tasks;

    // async equivalents to relevant methods in the reader
    // NOTE: changes should be made to the synchronous variants first, then ported here
    public partial class CsvReader
    {
        /// <summary>
        /// Asynchronously attempts to skip a record in the data, and increments <see cref="RecordNumber"/> if successful.
        /// </summary>
        /// <returns>
        /// <see langword="true"/> if the record is successfully skipped, or <see langword="false"/> if there are no more records to skip.
        /// </returns>
        public async Task<bool> SkipRecordAsync()
        {
            return await this.SkipRecordsAsync(1, true).ConfigureAwait(false) == 1;
        }

        /// <summary>
        /// Asynchronously attempts to skip a record in the data, optionally incrementing <see cref="RecordNumber"/> if successful.
        /// </summary>
        /// <param name="incrementRecordNumber">
        /// <see langword="true"/> to increment <see cref="RecordNumber"/> upon successfully skipping a record.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the record is successfully skipped, or <see langword="false"/> if there are no more records to skip.
        /// </returns>
        public async Task<bool> SkipRecordAsync(bool incrementRecordNumber)
        {
            return await this.SkipRecordsAsync(1, incrementRecordNumber).ConfigureAwait(false) == 1;
        }

        /// <summary>
        /// Asynchronously attempts to skip <paramref name="count"/> records in the data, and increments <see cref="RecordNumber"/> by the number of records actually skipped.
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
        public async Task<int> SkipRecordsAsync(int count)
        {
            return await this.SkipRecordsAsync(count, true).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously attempts to skip <paramref name="count"/> records in the data, optionally incrementing <see cref="RecordNumber"/> by the number of records actually skipped.
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
        public async Task<int> SkipRecordsAsync(int count, bool incrementRecordNumber)
        {
            this.EnsureNotDisposed();
            var skipped = await this.parser.SkipRecordsAsync(count).ConfigureAwait(false);

            if (incrementRecordNumber)
            {
                this.recordNumber += skipped;
            }

            return skipped;
        }

        /// <summary>
        /// Asynchronously reads the first record from the underlying CSV data and assigns it to <see cref="HeaderRecord"/>.
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
        public async Task<HeaderRecord> ReadHeaderRecordAsync()
        {
            this.EnsureNotDisposed();
            this.EnsureNotPassedFirstRecord();

            if (await this.parser.ParseRecordsAsync(null, this.buffer, 0, 1).ConfigureAwait(false) == 1)
            {
                ++this.recordNumber;
                this.headerRecord = new HeaderRecord(this.buffer[0]);
                return this.headerRecord;
            }

            return null;
        }

        /// <summary>
        /// Asynchronously reads a <see cref="DataRecord"/> from the underlying CSV.
        /// </summary>
        /// <returns>
        /// The <see cref="DataRecord"/> that was read, or <see langword="null"/> if there are no more records to read.
        /// </returns>
        public async Task<DataRecord> ReadDataRecordAsync()
        {
            this.EnsureNotDisposed();

            if (await this.parser.ParseRecordsAsync(this.headerRecord, this.buffer, 0, 1).ConfigureAwait(false) == 1)
            {
                ++this.recordNumber;
                return this.buffer[0];
            }

            return null;
        }

        /// <summary>
        /// Asynchronously reads at most <paramref name="count"/> <see cref="DataRecord"/>s and populates <paramref name="buffer"/> with them, beginning at index <paramref name="offset"/>.
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
        public async Task<int> ReadDataRecordsAsync(DataRecord[] buffer, int offset, int count)
        {
            this.EnsureNotDisposed();

            var read = await this.parser.ParseRecordsAsync(this.headerRecord, buffer, offset, count).ConfigureAwait(false);
            this.recordNumber += read;
            return read;
        }
    }
}