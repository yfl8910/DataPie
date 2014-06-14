namespace Kent.Boogaart.KBCsv.Internal
{
    using System;
    using System.Collections.Generic;

    // provides a read-only wrapper around a Dictionary<TKey, TValue>, which the BCL does not have until 4.5
    internal sealed class ReadOnlyDictionary<TKey, TValue> : IDictionary<TKey, TValue>
    {
        private readonly IDictionary<TKey, TValue> underlyingDictionary;

        public ReadOnlyDictionary(IDictionary<TKey, TValue> underlyingDictionary)
        {
            //underlyingDictionary.AssertNotNull("underlyingDictionary");
            this.underlyingDictionary = underlyingDictionary;
        }

        public int Count
        {
            get { return this.underlyingDictionary.Count; }
        }

        public bool IsReadOnly
        {
            get { return true; }
        }

        public ICollection<TKey> Keys
        {
            get { return this.underlyingDictionary.Keys; }
        }

        public ICollection<TValue> Values
        {
            get { return this.underlyingDictionary.Values; }
        }

        public TValue this[TKey key]
        {
            get { return this.underlyingDictionary[key]; }
            set { throw new NotSupportedException(); }
        }

        public void Add(TKey key, TValue value)
        {
            throw new NotSupportedException();
        }

        public void Add(KeyValuePair<TKey, TValue> item)
        {
            throw new NotSupportedException();
        }

        public void Clear()
        {
            throw new NotSupportedException();
        }

        public bool ContainsKey(TKey key)
        {
            return this.underlyingDictionary.ContainsKey(key);
        }

        public bool Remove(TKey key)
        {
            throw new NotSupportedException();
        }

        public bool TryGetValue(TKey key, out TValue value)
        {
            return this.underlyingDictionary.TryGetValue(key, out value);
        }

        public bool Contains(KeyValuePair<TKey, TValue> item)
        {
            return this.underlyingDictionary.Contains(item);
        }

        public void CopyTo(KeyValuePair<TKey, TValue>[] array, int arrayIndex)
        {
            this.underlyingDictionary.CopyTo(array, arrayIndex);
        }

        public bool Remove(KeyValuePair<TKey, TValue> item)
        {
            throw new NotSupportedException();
        }

        public IEnumerator<KeyValuePair<TKey, TValue>> GetEnumerator()
        {
            return this.underlyingDictionary.GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return this.underlyingDictionary.GetEnumerator();
        }
    }
}
