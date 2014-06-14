namespace Kent.Boogaart.KBCsv
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Diagnostics;
    using System.Text;
    using Kent.Boogaart.HelperTrinity.Extensions;

    /// <summary>
    /// A base class for CSV records.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A record consists of zero or more non-<see langword="null"/> <see cref="string"/> values, which are accessible via their index. <see cref="Count"/> returns the number of values in the record.
    /// A record may be read-only, as indicated by the <see cref="IsReadOnly"/> property. Read-only records will throw an exception if any attempt is made to modify the record,
    /// such as by calling <see cref="Add"/> or <see cref="Remove"/>.
    /// </para>
    /// </remarks>
    public abstract class RecordBase : IList<string>, IEquatable<RecordBase>
    {
        private const char ValueSeparator = (char)0x2022;
        private readonly IList<string> values;

        /// <summary>
        /// Initializes a new instance of the RecordBase class.
        /// </summary>
        /// <param name="readOnly">
        /// <see langword="true"/> if the record is read-only, otherwise <see langword="false"/>.
        /// </param>
        /// <param name="values">
        /// The values comprising the record.
        /// </param>
        protected RecordBase(bool readOnly, IEnumerable<string> values)
        {
            this.values = new List<string>(values);

            if (readOnly)
            {
                this.values = new ReadOnlyCollection<string>(this.values);
            }
        }

        // used internally by the parser to speed up the creation of parsed records
        internal RecordBase(IList<string> values)
        {
            Debug.Assert(values != null, "Expecting non-null values.");
            Debug.Assert(values.IsReadOnly, "Expecting only read-only values.");

            this.values = values;
        }

        /// <summary>
        /// Gets or sets a value in this record by index.
        /// </summary>
        /// <param name="index">
        /// The index of the value to retrieve.
        /// </param>
        /// <returns>
        /// The value at the specified index.
        /// </returns>
        public virtual string this[int index]
        {
            get
            {
                return this.values[index];
            }

            set
            {
                value.AssertNotNull("value");
                var oldValue = this.values[index];
                this.values[index] = value;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this record is read-only.
        /// </summary>
        public bool IsReadOnly
        {
            get { return this.values.IsReadOnly; }
        }

        /// <summary>
        /// Gets the number of values in this record.
        /// </summary>
        public int Count
        {
            get { return this.values.Count; }
        }

        /// <summary>
        /// Gets a value at the specified index, or <see langword="null"/> if the index is invalid.
        /// </summary>
        /// <param name="index">
        /// The index of the value to retrieve.
        /// </param>
        /// <returns>
        /// The value at the specified index, or <see langword="null"/> if the index is not valid.
        /// </returns>
        public string GetValueOrNull(int index)
        {
            if (index >= 0 && index < this.values.Count)
            {
                return this.values[index];
            }

            return null;
        }

        /// <summary>
        /// Clears all values in this record.
        /// </summary>
        public virtual void Clear()
        {
            this.values.Clear();
        }

        /// <summary>
        /// Determines whether this record contains the specified value.
        /// </summary>
        /// <param name="value">
        /// The value.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if this record contains <paramref name="value"/>, otherwise <see langword="false"/>.
        /// </returns>
        public bool Contains(string value)
        {
            value.AssertNotNull("value");
            return this.values.Contains(value);
        }

        /// <summary>
        /// Determines the index of the specified value within this record.
        /// </summary>
        /// <param name="value">
        /// The value.
        /// </param>
        /// <returns>
        /// The index of the value, or <c>-1</c> if the value does not exist in this record.
        /// </returns>
        public int IndexOf(string value)
        {
            return this.values.IndexOf(value);
        }

        /// <summary>
        /// Adds the specified value to this record.
        /// </summary>
        /// <param name="value">
        /// The value.
        /// </param>
        public virtual void Add(string value)
        {
            value.AssertNotNull("value");
            this.values.Add(value);
        }

        /// <summary>
        /// Inserts the specified value into this record.
        /// </summary>
        /// <param name="index">
        /// The index at which to insert the value.
        /// </param>
        /// <param name="value">
        /// The value to insert.
        /// </param>
        public virtual void Insert(int index, string value)
        {
            value.AssertNotNull("value");
            this.values.Insert(index, value);
        }

        /// <summary>
        /// Removes the specified value from this record.
        /// </summary>
        /// <param name="value">
        /// The value to remove.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the value was removed, otherwise <see langword="false"/>.
        /// </returns>
        public virtual bool Remove(string value)
        {
            value.AssertNotNull("value");
            return this.values.Remove(value);
        }

        /// <summary>
        /// Removes the value at the specified index.
        /// </summary>
        /// <param name="index">
        /// The index of the value to remove.
        /// </param>
        public virtual void RemoveAt(int index)
        {
            var value = this.values[index];
            this.values.RemoveAt(index);
        }

        /// <summary>
        /// Copies the values in this record to the specified array.
        /// </summary>
        /// <param name="array">
        /// The array.
        /// </param>
        /// <param name="arrayIndex">
        /// The starting index in the array to which values will be copied.
        /// </param>
        public void CopyTo(string[] array, int arrayIndex)
        {
            this.values.CopyTo(array, arrayIndex);
        }

        /// <summary>
        /// Determines whether this record is equal to another.
        /// </summary>
        /// <remarks>
        /// Records are considered equal if they have the same number of values, and every corresponding value is equal.
        /// </remarks>
        /// <param name="other">
        /// The other record.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if the records are equal, otherwise <see langword="false"/>.
        /// </returns>
        public bool Equals(RecordBase other)
        {
            if (other == null)
            {
                return false;
            }

            if (object.ReferenceEquals(other, this))
            {
                return true;
            }

            if (values.Count != other.values.Count)
            {
                return false;
            }

            for (var i = 0; i < values.Count; ++i)
            {
                if (!string.Equals(values[i], other.values[i], StringComparison.CurrentCulture))
                {
                    return false;
                }
            }

            return true;
        }

        /// <inheritdoc/>
        public override bool Equals(object obj)
        {
            return this.Equals(obj as RecordBase);
        }

        /// <inheritdoc/>
        public override int GetHashCode()
        {
            var hash = 17;

            for (var i = 0; i < values.Count; ++i)
            {
                hash = hash * 23 + values[i].GetHashCode();
            }

            return hash;
        }

        /// <inheritdoc/>
        public sealed override string ToString()
        {
            var retVal = new StringBuilder();

            foreach (var val in values)
            {
                retVal.Append(val).Append(ValueSeparator);
            }

            return retVal.ToString();
        }

        /// <inheritdoc/>
        IEnumerator<string> IEnumerable<string>.GetEnumerator()
        {
            return this.values.GetEnumerator();
        }

        /// <inheritdoc/>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.values.GetEnumerator();
        }
    }
}