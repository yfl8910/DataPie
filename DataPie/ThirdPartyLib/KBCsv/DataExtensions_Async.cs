namespace Kent.Boogaart.KBCsv.Extensions.Data
{
    using Kent.Boogaart.HelperTrinity.Extensions;
    using System;
    using System.Data;
    using System.Threading.Tasks;

    // async equivalents to data extension methods
    // NOTE: changes should be made to the synchronous variants first, then ported here
    public static partial class DataExtensions
    {
        /// <summary>
        /// Creates a table in <paramref name="this"/> and populates it with data read asynchronously from <paramref name="csvReader"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// The name of the table created and added to <paramref name="this"/> is <see cref="DefaultTableName"/>. All records from <paramref name="csvReader"/> will be read and added to the table.
        /// </para>
        /// <para>
        /// <paramref name="csvReader"/> must have a <see cref="HeaderRecord"/>, which is used to populate the column names of the <see cref="DataTable"/>.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataSet"/>.
        /// </param>
        /// <param name="csvReader">
        /// The <see cref="CsvReader"/>.
        /// </param>
        /// <returns>
        /// The number of rows added to the <see cref="DataTable"/> (and therefore the number of data records read from <paramref name="csvReader"/>).
        /// </returns>
        public async static Task<int> FillAsync(this DataSet @this, CsvReader csvReader)
        {
            return await @this.FillAsync(csvReader, DefaultTableName).ConfigureAwait(false);
        }

        /// <summary>
        /// Creates a table in <paramref name="this"/> and populates it with data read asynchronously from <paramref name="csvReader"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// All records from <paramref name="csvReader"/> will be read and added to the table.
        /// </para>
        /// <para>
        /// <paramref name="csvReader"/> must have a <see cref="HeaderRecord"/>, which is used to populate the column names of the <see cref="DataTable"/>.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataSet"/>.
        /// </param>
        /// <param name="csvReader">
        /// The <see cref="CsvReader"/>.
        /// </param>
        /// <param name="tableName">
        /// The name of the table to create and add to <paramref name="this"/>
        /// </param>
        /// <returns>
        /// The number of rows added to the <see cref="DataTable"/> (and therefore the number of data records read from <paramref name="csvReader"/>).
        /// </returns>
        public async static Task<int> FillAsync(this DataSet @this, CsvReader csvReader, string tableName)
        {
            return await @this.FillAsync(csvReader, tableName, null).ConfigureAwait(false);
        }

        /// <summary>
        /// Creates a table in <paramref name="this"/> and populates it with data read asynchronously from <paramref name="csvReader"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// <paramref name="csvReader"/> must have a <see cref="HeaderRecord"/>, which is used to populate the column names of the <see cref="DataTable"/>.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataSet"/>.
        /// </param>
        /// <param name="csvReader">
        /// The <see cref="CsvReader"/>.
        /// </param>
        /// <param name="tableName">
        /// The name of the table to create and add to <paramref name="this"/>
        /// </param>
        /// <param name="maximumRecords">
        /// The maximum number of records to read and add to the <see cref="DataTable"/>.
        /// </param>
        /// <returns>
        /// The number of rows added to the <see cref="DataTable"/> (and therefore the number of data records read from <paramref name="csvReader"/>).
        /// </returns>
        public async static Task<int> FillAsync(this DataSet @this, CsvReader csvReader, string tableName, int? maximumRecords)
        {
            @this.AssertNotNull("@this");
            tableName.AssertNotNull("tableName");

            var table = @this.Tables.Add(tableName);
            return await table.FillAsync(csvReader, maximumRecords).ConfigureAwait(false);
        }

        /// <summary>
        /// Populates <paramref name="this"/> with data read asynchronously from <paramref name="csvReader"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// All records from <paramref name="csvReader"/> will be read and added to the table.
        /// </para>
        /// <para>
        /// If <paramref name="this"/> has columns defined, those columns will be used when populating the data. If no columns have been defined, <paramref name="csvReader"/> must have a
        /// <see cref="HeaderRecord"/>, which is then used to define the columns for <paramref name="this"/>. If any data record has more values than can fit into the columns defined on
        /// <paramref name="this"/>, an exception is thrown.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataTable"/>.
        /// </param>
        /// <param name="csvReader">
        /// The <see cref="CsvReader"/>.
        /// </param>
        /// <returns>
        /// The number of rows added to <paramref name="this"/> (and therefore the number of data records read from <paramref name="csvReader"/>).
        /// </returns>
        public async static Task<int> FillAsync(this DataTable @this, CsvReader csvReader)
        {
            return await @this.FillAsync(csvReader, null).ConfigureAwait(false);
        }

        /// <summary>
        /// Populates <paramref name="this"/> with data read asynchronously from <paramref name="csvReader"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// If <paramref name="this"/> has columns defined, those columns will be used when populating the data. If no columns have been defined, <paramref name="csvReader"/> must have a
        /// <see cref="HeaderRecord"/>, which is then used to define the columns for <paramref name="this"/>. If any data record has more values than can fit into the columns defined on
        /// <paramref name="this"/>, an exception is thrown.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataTable"/>.
        /// </param>
        /// <param name="csvReader">
        /// The <see cref="CsvReader"/>.
        /// </param>
        /// <param name="maximumRecords">
        /// The maximum number of records to read and add to <paramref name="this"/>.
        /// </param>
        /// <returns>
        /// The number of rows added to <paramref name="this"/> (and therefore the number of data records read from <paramref name="csvReader"/>).
        /// </returns>
        public async static Task<int> FillAsync(this DataTable @this, CsvReader csvReader, int? maximumRecords)
        {
            @this.AssertNotNull("@this");
            csvReader.AssertNotNull("csvReader");
            exceptionHelper.ResolveAndThrowIf(maximumRecords.GetValueOrDefault() < 0, "maximumRecordsMustBePositive");

            if (@this.Columns.Count == 0)
            {
                // table has no columns, so we need to use the CSV header record to populate them
                exceptionHelper.ResolveAndThrowIf(csvReader.HeaderRecord == null, "noColumnsAndNoHeaderRecord");

                foreach (var columnName in csvReader.HeaderRecord)
                {
                    @this.Columns.Add(columnName);
                }
            }

            var remaining = maximumRecords.GetValueOrDefault(int.MaxValue);
            var buffer = new DataRecord[16];

            while (remaining > 0)
            {
                var read = await csvReader.ReadDataRecordsAsync(buffer, 0, Math.Min(buffer.Length, remaining)).ConfigureAwait(false);

                if (read == 0)
                {
                    // no more data
                    break;
                }

                for (var i = 0; i < read; ++i)
                {
                    var record = buffer[i];
                    exceptionHelper.ResolveAndThrowIf(record.Count > @this.Columns.Count, "moreValuesThanColumns", @this.Columns.Count, record.Count);

                    var recordAsStrings = new string[record.Count];
                    record.CopyTo(recordAsStrings, 0);
                    @this.Rows.Add(recordAsStrings);
                }

                remaining -= read;
            }

            return maximumRecords.GetValueOrDefault(int.MaxValue) - remaining;
        }

        /// <summary>
        /// Asynchronously writes all rows in <paramref name="this"/> to <paramref name="csvWriter"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// All rows in <paramref name="this"/> will be written to <paramref name="csvWriter"/>. A header record will also be written, which will be comprised of the column names defined for
        /// <paramref name="this"/>.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataTable"/>.
        /// </param>
        /// <param name="csvWriter">
        /// The <see cref="CsvWriter"/>.
        /// </param>
        /// <returns>
        /// The actual number of rows from <paramref name="this"/> written to <paramref name="csvWriter"/>.
        /// </returns>
        public async static Task<int> WriteCsvAsync(this DataTable @this, CsvWriter csvWriter)
        {
            return await @this.WriteCsvAsync(csvWriter, true).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes all rows in <paramref name="this"/> to <paramref name="csvWriter"/>.
        /// </summary>
        /// <remarks>
        /// <para>
        /// All rows in <paramref name="this"/> will be written to <paramref name="csvWriter"/>.
        /// </para>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataTable"/>.
        /// </param>
        /// <param name="csvWriter">
        /// The <see cref="CsvWriter"/>.
        /// </param>
        /// <param name="writeHeaderRecord">
        /// If <see langword="true"/>, a header record will also be written, which will be comprised of the column names defined for <paramref name="this"/>.
        /// </param>
        /// <returns>
        /// The actual number of rows from <paramref name="this"/> written to <paramref name="csvWriter"/>.
        /// </returns>
        public async static Task<int> WriteCsvAsync(this DataTable @this, CsvWriter csvWriter, bool writeHeaderRecord)
        {
            return await @this.WriteCsvAsync(csvWriter, writeHeaderRecord, null).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes all rows in <paramref name="this"/> to <paramref name="csvWriter"/>.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataTable"/>.
        /// </param>
        /// <param name="csvWriter">
        /// The <see cref="CsvWriter"/>.
        /// </param>
        /// <param name="writeHeaderRecord">
        /// If <see langword="true"/>, a header record will also be written, which will be comprised of the column names defined for <paramref name="this"/>.
        /// </param>
        /// <param name="maximumRows">
        /// The maximum number of rows from <paramref name="this"/> that should be written to <paramref name="csvWriter"/>.
        /// </param>
        /// <returns>
        /// The actual number of rows from <paramref name="this"/> written to <paramref name="csvWriter"/>.
        /// </returns>
        public async static Task<int> WriteCsvAsync(this DataTable @this, CsvWriter csvWriter, bool writeHeaderRecord, int? maximumRows)
        {
            return await @this.WriteCsvAsync(csvWriter, writeHeaderRecord, maximumRows, o => o == null ? string.Empty : o.ToString()).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes all rows in <paramref name="this"/> to <paramref name="csvWriter"/>.
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="this">
        /// The <see cref="DataTable"/>.
        /// </param>
        /// <param name="csvWriter">
        /// The <see cref="CsvWriter"/>.
        /// </param>
        /// <param name="writeHeaderRecord">
        /// If <see langword="true"/>, a header record will also be written, which will be comprised of the column names defined for <paramref name="this"/>.
        /// </param>
        /// <param name="maximumRows">
        /// The maximum number of rows from <paramref name="this"/> that should be written to <paramref name="csvWriter"/>.
        /// </param>
        /// <param name="objectToStringConverter">
        /// Provides a means of converting values in the <see cref="DataRow"/>s to <see cref="String"/>s.
        /// </param>
        /// <returns>
        /// The actual number of rows from <paramref name="this"/> written to <paramref name="csvWriter"/>.
        /// </returns>
        public async static Task<int> WriteCsvAsync(this DataTable @this, CsvWriter csvWriter, bool writeHeaderRecord, int? maximumRows, Func<object, string> objectToStringConverter)
        {
            @this.AssertNotNull("@this");
            csvWriter.AssertNotNull("csvWriter");
            objectToStringConverter.AssertNotNull("objectToStringConverter");

            var num = 0;

            if (writeHeaderRecord)
            {
                var columnNames = new string[@this.Columns.Count];

                for (var i = 0; i < columnNames.Length; ++i)
                {
                    columnNames[i] = @this.Columns[i].ColumnName;
                }

                await csvWriter.WriteRecordAsync(columnNames).ConfigureAwait(false);
            }

            var maximum = maximumRows.GetValueOrDefault(int.MaxValue);
            var buffer = new DataRecord[16];
            var bufferOffset = 0;

            foreach (DataRow row in @this.Rows)
            {
                var record = new DataRecord();

                for (var i = 0; i < row.ItemArray.Length; ++i)
                {
                    record.Add(objectToStringConverter(row.ItemArray[i]));
                }

                buffer[bufferOffset++] = record;

                if (bufferOffset == buffer.Length)
                {
                    // buffer full
                    await csvWriter.WriteRecordsAsync(buffer, 0, buffer.Length).ConfigureAwait(false);
                    bufferOffset = 0;
                }

                if (++num == maximum)
                {
                    break;
                }
            }

            // write any outstanding data in buffer
            await csvWriter.WriteRecordsAsync(buffer, 0, bufferOffset).ConfigureAwait(false);

            return num;
        }
    }
}