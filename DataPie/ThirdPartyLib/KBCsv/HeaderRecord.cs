namespace Kent.Boogaart.KBCsv
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using Kent.Boogaart.HelperTrinity;
    using Kent.Boogaart.HelperTrinity.Extensions;
    using Kent.Boogaart.KBCsv.Internal;

    /// <summary>
    /// Represents the header record of a CSV file.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A <c>HeaderRecord</c> is a specialized <see cref="RecordBase"/> whose values are the names of columns in the CSV. It permits the index of each column to be obtained given
    /// its name (see <see cref="P:Kent.Boogaart.KBCsv.HeaderRecord.Item(System.String)"/> and <see cref="GetColumnIndexOrNull"/>).
    /// </para>
    /// <para>
    /// Note that there is a cost to maintaining the indexes of columns by their name. If possible, favor passing in all column names during construction. Otherwise, favor using
    /// <see cref="Add"/> only, avoiding <see cref="Insert"/>, <see cref="Remove"/> and <see cref="RemoveAt"/> wherever possible.
    /// </para>
    /// </remarks>
    public sealed class HeaderRecord : RecordBase
    {
        private static readonly ExceptionHelper exceptionHelper = new ExceptionHelper(typeof(HeaderRecord));
        private readonly IDictionary<string, int> columnNameToIndexMap;

        /// <summary>
        /// Initializes a new instance of the HeaderRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant header record will have no values, but is not read-only.
        /// </remarks>
        public HeaderRecord()
            : this(false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the HeaderRecord class with the specified column names.
        /// </summary>
        /// <remarks>
        /// The resultant header record will have the specified column names as values and is not read-only.
        /// </remarks>
        /// <param name="columnNames">
        /// The names of the columns in the header record.
        /// </param>
        public HeaderRecord(params string[] columnNames)
            : this(false, columnNames)
        {
        }

        /// <summary>
        /// Initializes a new instance of the HeaderRecord class.
        /// </summary>
        /// <param name="columnNames">
        /// The names of the columns in the header record.
        /// </param>
        /// <param name="readOnly">
        /// <see langword="true"/> if the header record is read-only, otherwise <see langword="false"/>.
        /// </param>
        public HeaderRecord(bool readOnly, params string[] columnNames)
            : this(readOnly, (IEnumerable<string>)columnNames)
        {
        }

        /// <summary>
        /// Initializes a new instance of the HeaderRecord class with the specified column names.
        /// </summary>
        /// <remarks>
        /// The resultant header record will have the specified column names as values and is not read-only.
        /// </remarks>
        /// <param name="columnNames">
        /// The names of the columns in the header record.
        /// </param>
        public HeaderRecord(IEnumerable<string> columnNames)
            : this(false, columnNames)
        {
        }

        /// <summary>
        /// Initializes a new instance of the HeaderRecord class.
        /// </summary>
        /// <param name="columnNames">
        /// The names of the columns in the header record.
        /// </param>
        /// <param name="readOnly">
        /// <see langword="true"/> if the header record is read-only, otherwise <see langword="false"/>.
        /// </param>
        public HeaderRecord(bool readOnly, IEnumerable<string> columnNames)
            : base(readOnly, columnNames)
        {
            columnNames.AssertNotNull("columnNames", true);

            this.columnNameToIndexMap = new Dictionary<string, int>();
            this.PopulateColumnNameToIndexMap(0, true);

            if (readOnly)
            {
                this.columnNameToIndexMap = new ReadOnlyDictionary<string, int>(this.columnNameToIndexMap);
            }
        }

        // used internally by the parser to speed up the creation of parsed records
        internal HeaderRecord(IList<string> columnNames)
            : base(columnNames)
        {
            Debug.Assert(this.IsReadOnly, "Expecting to be read-only.");

            this.columnNameToIndexMap = new Dictionary<string, int>();
            this.PopulateColumnNameToIndexMap(0, true);
            this.columnNameToIndexMap = new ReadOnlyDictionary<string, int>(this.columnNameToIndexMap);
        }

        /// <summary>
        /// Gets the index of the column with the specified name.
        /// </summary>
        /// <param name="columnName">
        /// The column name.
        /// </param>
        /// <returns>
        /// The index of the column with the specified name.
        /// </returns>
        public int this[string columnName]
        {
            get
            {
                try
                {
                    return this.columnNameToIndexMap[columnName];
                }
                catch (KeyNotFoundException ex)
                {
                    throw exceptionHelper.Resolve("columnNotFound", ex, columnName);
                }
            }
        }

        /// <summary>
        /// Gets the index of the specified column, or <see langword="null"/> if the column does not exist in this header record.
        /// </summary>
        /// <param name="columnName">
        /// The column name.
        /// </param>
        /// <returns>
        /// The index of the column, or <see langword="null"/> if the column does not exist in this header record.
        /// </returns>
        public int? GetColumnIndexOrNull(string columnName)
        {
            int index;

            if (this.columnNameToIndexMap.TryGetValue(columnName, out index))
            {
                return index;
            }

            return null;
        }

        /// <inheritdoc/>
        public override string this[int index]
        {
            get
            {
                return base[index];
            }

            set
            {
                this.columnNameToIndexMap.Remove(this[index]);
                base[index] = value;
                this.columnNameToIndexMap[value] = index;
            }
        }

        /// <inheritdoc/>
        public override void Add(string value)
        {
            exceptionHelper.ResolveAndThrowIf(this.Any(x => string.Equals(value, x, StringComparison.CurrentCulture)), "duplicateValue", value);
            base.Add(value);
            this.columnNameToIndexMap[value] = this.Count - 1;
        }

        /// <inheritdoc/>
        public override void Insert(int index, string value)
        {
            exceptionHelper.ResolveAndThrowIf(this.Any(x => string.Equals(value, x, StringComparison.CurrentCulture)), "duplicateValue", value);
            base.Insert(index, value);
            this.PopulateColumnNameToIndexMap(index, false);
        }

        /// <inheritdoc/>
        public override void Clear()
        {
            base.Clear();
            this.columnNameToIndexMap.Clear();
        }

        /// <inheritdoc/>
        public override bool Remove(string value)
        {
            var columnIndex = this.GetColumnIndexOrNull(value);

            if (columnIndex.HasValue)
            {
                // we use RemoveAt here instead so that our index population is more efficient
                this.RemoveAt(columnIndex.GetValueOrDefault());
                return true;
            }

            return false;
        }

        /// <inheritdoc/>
        public override void RemoveAt(int index)
        {
            base.RemoveAt(index);
            this.PopulateColumnNameToIndexMap(index, false);
        }

        private void PopulateColumnNameToIndexMap(int startIndex, bool checkForDuplicates)
        {
            for (var index = startIndex; index < this.Count; ++index)
            {
                var columnName = this[index];

                if (checkForDuplicates && this.columnNameToIndexMap.ContainsKey(columnName))
                {
                    throw exceptionHelper.Resolve("duplicateValue", columnName);
                }

                this.columnNameToIndexMap[columnName] = index;
            }
        }
    }
}
