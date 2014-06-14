namespace Kent.Boogaart.KBCsv
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Kent.Boogaart.HelperTrinity;

    /// <summary>
    /// Represents a data record in a CSV file.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A <c>DataRecord</c> represents a CSV record that is not the header record. Values in the data record can be accessed by their index. A data record can have an associated
    /// <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> (exposed via <see cref="HeaderRecord"/>), in which case values in the data record may also be accessed via a column name.
    /// </para>
    /// </remarks>
    public sealed class DataRecord : RecordBase, IEquatable<DataRecord>
    {
        private static readonly ExceptionHelper exceptionHelper = new ExceptionHelper(typeof(DataRecord));
        private readonly HeaderRecord headerRecord;

        /// <summary>
        /// Initializes a new instance of the DataRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant data record will have no values, but is not read-only.
        /// </remarks>
        public DataRecord()
            : this(null, false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the DataRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant data record will have no values, but is not read-only. It will use the specified <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> (which will therefore
        /// be returned from <see cref="HeaderRecord"/>).
        /// </remarks>
        /// <param name="headerRecord">
        /// An optional <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> associated with this <c>DataRecord</c>.
        /// </param>
        public DataRecord(HeaderRecord headerRecord)
            : this(headerRecord, false)
        {
        }

        /// <summary>
        /// Initializes a new instance of the DataRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant data record will the specified values, and is not read-only. It will use the specified <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> (which will therefore
        /// be returned from <see cref="HeaderRecord"/>).
        /// </remarks>
        /// <param name="headerRecord">
        /// An optional <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> associated with this <c>DataRecord</c>.
        /// </param>
        /// <param name="values">
        /// The values comprising this <c>DataRecord</c>.
        /// </param>
        public DataRecord(HeaderRecord headerRecord, params string[] values)
            : this(headerRecord, false, values)
        {
        }

        /// <summary>
        /// Initializes a new instance of the DataRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant data record will have the specified values, and may or may not be read-only. It will use the specified <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> (which will therefore
        /// be returned from <see cref="HeaderRecord"/>).
        /// </remarks>
        /// <param name="headerRecord">
        /// An optional <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> associated with this <c>DataRecord</c>.
        /// </param>
        /// <param name="readOnly">
        /// <see langword="true"/> to mark this <c>DataRecord</c> as read-only.
        /// </param>
        /// <param name="values">
        /// The values comprising this <c>DataRecord</c>.
        /// </param>
        public DataRecord(HeaderRecord headerRecord, bool readOnly, params string[] values)
            : this(headerRecord, readOnly, (IEnumerable<string>)values)
        {
        }

        /// <summary>
        /// Initializes a new instance of the DataRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant data record will the specified values, and is not read-only. It will use the specified <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> (which will therefore
        /// be returned from <see cref="HeaderRecord"/>).
        /// </remarks>
        /// <param name="headerRecord">
        /// An optional <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> associated with this <c>DataRecord</c>.
        /// </param>
        /// <param name="values">
        /// The values comprising this <c>DataRecord</c>.
        /// </param>
        public DataRecord(HeaderRecord headerRecord, IEnumerable<string> values)
            : this(headerRecord, false, values)
        {
        }

        /// <summary>
        /// Initializes a new instance of the DataRecord class.
        /// </summary>
        /// <remarks>
        /// The resultant data record will have the specified values, and may or may not be read-only. It will use the specified <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> (which will therefore
        /// be returned from <see cref="HeaderRecord"/>).
        /// </remarks>
        /// <param name="headerRecord">
        /// An optional <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> associated with this <c>DataRecord</c>.
        /// </param>
        /// <param name="readOnly">
        /// <see langword="true"/> to mark this <c>DataRecord</c> as read-only.
        /// </param>
        /// <param name="values">
        /// The values comprising this <c>DataRecord</c>.
        /// </param>
        public DataRecord(HeaderRecord headerRecord, bool readOnly, IEnumerable<string> values)
            : base(readOnly, values)
        {
            this.headerRecord = headerRecord;
        }

        // used internally by the parser to speed up the creation of parsed records
        internal DataRecord(HeaderRecord headerRecord, IList<string> values)
            : base(values)
        {
            this.headerRecord = headerRecord;
        }

        /// <summary>
        /// Gets the <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> in use by this data record.
        /// </summary>
        /// <remarks>
        /// In order to get or set values in a data record via their column name, a header record must be provided. This property gets the <see cref="Kent.Boogaart.KBCsv.HeaderRecord"/> that
        /// this data record uses to resolve column indexes given a column name.
        /// </remarks>
        public HeaderRecord HeaderRecord
        {
            get { return this.headerRecord; }
        }

        /// <summary>
        /// Gets or sets a value in this data record by its column name.
        /// </summary>
        /// <param name="columnName">
        /// The name of the column.
        /// </param>
        /// <returns>
        /// The value in the specified column.
        /// </returns>
        public string this[string columnName]
        {
            get
            {
                exceptionHelper.ResolveAndThrowIf(this.headerRecord == null, "noHeader");
                return this[this.headerRecord[columnName]];
            }

            set
            {
                exceptionHelper.ResolveAndThrowIf(this.headerRecord == null, "noHeader");
                this[this.headerRecord[columnName]] = value;
            }
        }

        /// <summary>
        /// Gets a value in this data record given the column name, or <see langword="null"/> if the specified column does not exist, or if there is no value in this data record at the column's index.
        /// </summary>
        /// <param name="columnName">
        /// The name of the column.
        /// </param>
        /// <returns>
        /// The value in the specified column, or <see langword="null"/> if the column does not exist, or if there is no value in this data record at the column's index.
        /// </returns>
        public string GetValueOrNull(string columnName)
        {
            exceptionHelper.ResolveAndThrowIf(this.headerRecord == null, "noHeader");
            var columnIndex = this.headerRecord.GetColumnIndexOrNull(columnName);

            if (columnIndex.HasValue)
            {
                // GetValueOrDefault is a performance tweak: since we know there's a value we can use it safely
                return this.GetValueOrNull(columnIndex.GetValueOrDefault());
            }

            return null;
        }

        /// <summary>
        /// Determines whether this data record is equal to another.
        /// </summary>
        /// <remarks>
        /// Data records are considered equal if their <see cref="HeaderRecord"/>s are equal (or both absent), and if their values are equal.
        /// </remarks>
        /// <param name="other">
        /// The other data record.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the data records are equal, otherwise <see langword="false"/>.
        /// </returns>
        public bool Equals(DataRecord other)
        {
            if (other == null)
            {
                return false;
            }

            if (object.ReferenceEquals(other, this))
            {
                return true;
            }

            // this checks that all values are equal
            if (!base.Equals(other))
            {
                return false;
            }

            if ((this.headerRecord == null) && (other.headerRecord == null))
            {
                // both have no header and equal values, therefore they are equal
                return true;
            }
            else if (this.headerRecord != null && other.headerRecord != null && this.headerRecord.Equals(other.headerRecord))
            {
                // both have equal headers and equal values, therefore they are equal
                return true;
            }

            return false;
        }

        /// <inheritdoc/>
        public override bool Equals(object obj)
        {
            return this.Equals(obj as DataRecord);
        }

        /// <inheritdoc/>
        public override int GetHashCode()
        {
            var hash = base.GetHashCode();

            if (this.headerRecord != null)
            {
                hash = hash * 23 + this.headerRecord.GetHashCode();
            }

            return hash;
        }
    }
}
