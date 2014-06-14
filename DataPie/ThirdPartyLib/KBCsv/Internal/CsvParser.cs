namespace Kent.Boogaart.KBCsv.Internal
{
    using Kent.Boogaart.HelperTrinity;
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.IO;

    // implements the actual parsing logic, given a TextReader from which to read characters.
    internal sealed partial class CsvParser
    {
        public const int BufferSize = 4096;
        private static readonly ExceptionHelper exceptionHelper = new ExceptionHelper(typeof(CsvParser));
        private readonly TextReader reader;
        private readonly char[] buffer;
        private readonly ValueList values;
        private readonly ValueBuilder valueBuilder;
        private int bufferEndIndex;
        private int bufferIndex;
        private int specialCharacterMask;
        private int activeSpecialCharacterMask;
        private bool preserveLeadingWhiteSpace;
        private bool preserveTrailingWhiteSpace;
        private char valueSeparator;
        private char? valueDelimiter;

        public CsvParser(TextReader reader)
        {
            Debug.Assert(reader != null, "Reader must never be null.");

            this.reader = reader;
            this.buffer = new char[BufferSize];
            this.values = new ValueList();
            this.valueBuilder = new ValueBuilder(this);
            this.valueSeparator = Constants.DefaultValueSeparator;
            this.valueDelimiter = Constants.DefaultValueDelimiter;

            this.UpdateSpecialCharacterMask();
        }

        public TextReader TextReader
        {
            get { return this.reader; }
        }

        public bool PreserveLeadingWhiteSpace
        {
            get { return this.preserveLeadingWhiteSpace; }
            set { this.preserveLeadingWhiteSpace = value; }
        }

        public bool PreserveTrailingWhiteSpace
        {
            get { return this.preserveTrailingWhiteSpace; }
            set { this.preserveTrailingWhiteSpace = value; }
        }

        public char ValueSeparator
        {
            get { return this.valueSeparator; }
            set
            {
                exceptionHelper.ResolveAndThrowIf(value == this.valueDelimiter, "valueSeparatorAndDelimiterCannotMatch");

                this.valueSeparator = value;
                this.UpdateSpecialCharacterMask();
            }
        }

        public char? ValueDelimiter
        {
            get { return this.valueDelimiter; }
            set
            {
                exceptionHelper.ResolveAndThrowIf(value == this.valueSeparator, "valueSeparatorAndDelimiterCannotMatch");

                this.valueDelimiter = value;
                this.UpdateSpecialCharacterMask();
            }
        }

        public bool HasMoreRecords
        {
            get
            {
                if (this.bufferIndex < this.bufferEndIndex)
                {
                    // the buffer isn't empty so there must be more records
                    return true;
                }

                // the buffer is empty, so we only have more records if we successfully fill it
                return this.FillBuffer();
            }
        }

        private bool IsBufferEmpty
        {
            get { return this.bufferIndex == this.bufferEndIndex; }
        }

        public int SkipRecords(int skip)
        {
            // Performance Notes:
            //   - Checking !HasMoreRecords and exiting early actually degrades performance
            //     Looking at the IL, I assume this is because the common case (more records) results in a branch, whereas the uncommon case (no more records) does not

            var skipped = 0;
            var delimited = false;

            while (skipped < skip)
            {
                if (this.HasMoreRecords)
                {
                    while (true)
                    {
                        if (!this.IsBufferEmpty)
                        {
                            var ch = this.buffer[this.bufferIndex++];

                            if (!this.IsPossiblySpecialCharacter(ch))
                            {
                                // if it's definitely not a special character, then we can just continue on with the loop
                                continue;
                            }

                            if (!delimited)
                            {
                                if (ch == this.valueDelimiter)
                                {
                                    delimited = true;

                                    // since we're in a delimited area, the only special character is the value delimiter
                                    this.activeSpecialCharacterMask = this.valueDelimiter.Value;
                                }
                                else if (ch == Constants.CR)
                                {
                                    // we need to look at the next character, so make sure it is available
                                    if (this.IsBufferEmpty && !this.FillBufferWithoutNotify())
                                    {
                                        // last character available was CR, so we know we're done at this point
                                        break;
                                    }

                                    // we deal with CRLF right here by checking if the next character is LF, in which case we just discard it
                                    if (this.buffer[this.bufferIndex] == Constants.LF)
                                    {
                                        ++this.bufferIndex;
                                    }

                                    break;
                                }
                                else if (ch == Constants.LF)
                                {
                                    break;
                                }
                            }
                            else if (ch == this.valueDelimiter)
                            {
                                // we need to look at the next character, so make sure it is available
                                if (this.IsBufferEmpty && !this.FillBufferWithoutNotify())
                                {
                                    break;
                                }

                                if (this.buffer[this.bufferIndex] == this.valueDelimiter)
                                {
                                    // delimiter is escaped, so just swallow it
                                    ++this.bufferIndex;
                                }
                                else
                                {
                                    // delimiter isn't escaped, so we are no longer in a delimited area
                                    delimited = false;
                                    this.activeSpecialCharacterMask = this.specialCharacterMask;
                                }
                            }
                        }
                        else if (!this.FillBufferWithoutNotify())
                        {
                            // all out of data, so we successfully skipped the final record
                            break;
                        }
                    }

                    ++skipped;
                }
                else
                {
                    break;
                }
            }

            return skipped;
        }

        public int ParseRecords(HeaderRecord headerRecord, DataRecord[] buffer, int offset, int count)
        {
            // see performance notes in SkipRecord

            var ch = char.MinValue;
            var recordsParsed = 0;
            var delimited = false;

            for (var i = offset; i < offset + count; ++i)
            {
                while (true)
                {
                    if (!this.IsBufferEmpty)
                    {
                        ch = this.buffer[this.bufferIndex++];

                        if (!this.IsPossiblySpecialCharacter(ch))
                        {
                            // if it's definitely not a special character, then we can just append it and continue on with the loop
                            this.valueBuilder.NotifyPreviousCharIncluded(delimited);
                            continue;
                        }

                        if (!delimited)
                        {
                            if (ch == this.valueSeparator)
                            {
                                this.values.Add(this.valueBuilder.GetValueAndClear());
                            }
                            else if (ch == this.valueDelimiter)
                            {
                                this.valueBuilder.NotifyPreviousCharExcluded();
                                delimited = true;

                                // since we're in a delimited area, the only special character is the value delimiter
                                this.activeSpecialCharacterMask = this.valueDelimiter.Value;
                            }
                            else if (ch == Constants.CR)
                            {
                                // we need to look at the next character, so make sure it is available
                                if (this.IsBufferEmpty && !this.FillBuffer())
                                {
                                    // undelimited CR indicates the end of a record, so add the existing value and then exit
                                    buffer[i] = this.values.GetDataRecordAndClear(headerRecord, this.valueBuilder.GetValueAndClear());
                                    break;
                                }

                                // we deal with CRLF right here by checking if the next character is LF, in which case we just discard it
                                if (this.buffer[this.bufferIndex] == Constants.LF)
                                {
                                    ++this.bufferIndex;
                                }

                                // undelimited CR or CRLF both indicate the end of a record, so add the existing value and then exit
                                buffer[i] = this.values.GetDataRecordAndClear(headerRecord, this.valueBuilder.GetValueAndClear());
                                break;
                            }
                            else if (ch == Constants.LF)
                            {
                                // undelimited LF indicates the end of a record, so add the existing value and then exit
                                buffer[i] = this.values.GetDataRecordAndClear(headerRecord, this.valueBuilder.GetValueAndClear());
                                break;
                            }
                            else
                            {
                                // it wasn't a special character after all, so just append it
                                this.valueBuilder.NotifyPreviousCharIncluded(false);
                            }
                        }
                        else if (ch == this.valueDelimiter)
                        {
                            // we need to look at the next character, so make sure it is available
                            if (this.IsBufferEmpty && !this.FillBuffer())
                            {
                                // out of data
                                delimited = false;
                                this.activeSpecialCharacterMask = this.specialCharacterMask;
                                buffer[i] = this.values.GetDataRecordAndClear(headerRecord, this.valueBuilder.GetValueAndClear());
                                break;
                            }

                            if (this.buffer[this.bufferIndex] == this.valueDelimiter)
                            {
                                // delimiter is escaped, so append it to the value and discard the escape character
                                this.valueBuilder.NotifyPreviousCharExcluded();
                                ++this.bufferIndex;
                                this.valueBuilder.NotifyPreviousCharIncluded(true);
                            }
                            else
                            {
                                // delimiter isn't escaped, so we are no longer in a delimited area
                                this.valueBuilder.NotifyPreviousCharExcluded();
                                delimited = false;
                                this.activeSpecialCharacterMask = this.specialCharacterMask;
                            }
                        }
                        else
                        {
                            // it wasn't a special character after all, so just append it
                            this.valueBuilder.NotifyPreviousCharIncluded(true);
                        }
                    }
                    else if (!this.FillBuffer())
                    {
                        if (this.valueBuilder.HasValue)
                        {
                            // a value is outstanding, so add it
                            this.values.Add(this.valueBuilder.GetValueAndClear());
                        }

                        if (ch == this.valueSeparator)
                        {
                            // special case: last character is a separator, which means there should be an empty value after it. eg. "foo," results in ["foo", ""]
                            buffer[i] = this.values.GetDataRecordAndClear(headerRecord, string.Empty);
                            break;
                        }
                        else
                        {
                            var record = this.values.GetDataRecordAndClear(headerRecord);

                            if (record != null)
                            {
                                buffer[i] = record;
                                ++recordsParsed;
                            }
                        }

                        // data exhausted - we're done, even though we may not have filled the records array
                        return recordsParsed;
                    }
                }

                ++recordsParsed;
            }

            return recordsParsed;
        }

        private static bool IsWhiteSpace(char ch)
        {
            return ch == Constants.Space || ch == Constants.Tab;
        }

        // update the mask used to quickly detect characters that are definitely not special (effectively a bloom filter, where false positives are possible but not false negatives)
        // we have two copies of the mask so that it can be swapped around during parsing to speed things up
        private void UpdateSpecialCharacterMask()
        {
            this.specialCharacterMask = this.valueSeparator | this.valueDelimiter.GetValueOrDefault() | Constants.CR | Constants.LF;
            this.activeSpecialCharacterMask = this.specialCharacterMask;
        }

        // gets a value indicating whether the character is possibly a special character
        // as indicated by the name, false positives are possible, false negatives are not
        // that is, this may return true even for a character that isn't special, but will never return false for a character that is special
        private bool IsPossiblySpecialCharacter(char ch)
        {
            return (ch & this.activeSpecialCharacterMask) == ch;
        }

        // fill the character buffer with data from the text reader
        private bool FillBuffer()
        {
            Debug.Assert(this.IsBufferEmpty, "Buffer not empty.");

            this.valueBuilder.NotifyBufferRefilling();
            this.bufferEndIndex = this.reader.Read(this.buffer, 0, BufferSize);
            this.bufferIndex = 0;

            return this.bufferEndIndex > 0;
        }

        // fill the character buffer with data from the text reader. Does not notify the value builder that the fill is taking place, which is useful when the value builder is irrelevant (such as when skipping records)
        private bool FillBufferWithoutNotify()
        {
            Debug.Assert(this.IsBufferEmpty, "Buffer not empty.");

            this.bufferEndIndex = this.reader.Read(this.buffer, 0, BufferSize);
            this.bufferIndex = 0;

            return this.bufferEndIndex > 0;
        }

        // lightweight list to store the values comprising the record currently being parsed
        // using this out-performs List<string> significantly
        // aggressive inlining also has a significantly positive effect on performance
        private sealed class ValueList
        {
            private string[] values;
            private int valueEndIndex;

            public ValueList()
            {
                this.values = new string[64];
            }

            public void Add(string value)
            {
                this.EnsureSufficientCapacity();
                this.values[this.valueEndIndex++] = value;
            }

            // convert this value list to a data record (null if there are no values), and clear ready to construct the next list of values
            public DataRecord GetDataRecordAndClear(HeaderRecord headerRecord)
            {
                if (this.valueEndIndex == 0)
                {
                    return null;
                }

                var result = new string[this.valueEndIndex];
                Array.Copy(this.values, 0, result, 0, this.valueEndIndex);

                // clear
                this.valueEndIndex = 0;

                // cast is to select the internal constructor, which is faster
                return new DataRecord(headerRecord, (IList<string>)result);
            }

            // convert this value list to a data record, placing the extra value at the end of the record's values, then clear ready to construct the next list of values
            // the extra parameter saves the client code having to add the value and then call GetDataRecordAndClear, which is more expensive than just doing it in one step
            public DataRecord GetDataRecordAndClear(HeaderRecord headerRecord, string extra)
            {
                var result = new string[this.valueEndIndex + 1];
                Array.Copy(this.values, 0, result, 0, this.valueEndIndex);
                result[this.valueEndIndex] = extra;

                // clear
                this.valueEndIndex = 0;

                // cast is to select the internal constructor, which is faster
                return new DataRecord(headerRecord, (IList<string>)result);
            }

            private void EnsureSufficientCapacity()
            {
                if (this.valueEndIndex == this.values.Length)
                {
                    // need to reallocate larger values array
                    var newValues = new string[this.values.Length * 2];
                    Array.Copy(this.values, 0, newValues, 0, this.values.Length);
                    this.values = newValues;
                }
            }
        }

        // builds values on behalf of the parser
        // avoids copying data wherever possible, instead referring back to the parser's buffer
        private sealed class ValueBuilder
        {
            private readonly CsvParser parser;

            // the total length of the value being built, regardless of whether it is stored in the parser's buffer or our own local buffer
            private int length;

            // the index into the parser's buffer where the value (or value part) starts, along with the length of the value (or value part) in the parser's buffer
            private int bufferStartIndex;
            private int bufferLength;

            // a local buffer (only used if necessary), along with the length of the value (or value part) stored within it
            private char[] localBuffer;
            private int localBufferLength;

            // indexes (relative to the start of the value) indicating where the first and last delimited characters are, if at all
            private int? delimitedStartIndex;
            private int? delimitedEndIndex;

            public ValueBuilder(CsvParser parser)
            {
                this.parser = parser;

                // to make the resize logic faster, our local buffer is the same size as the parser's buffer
                this.localBuffer = new char[BufferSize];
            }

            public bool HasValue
            {
                get { return this.length > 0; }
            }

            // tell the value builder that the previous character parsed should be included in the value
            // the delimited parameter tells the value builder whether the character in question appeared within a delimited area
            public void NotifyPreviousCharIncluded(bool delimited)
            {
                if (this.bufferLength == 0)
                {
                    // this is the first included char (for this demarcation), so we need to set our buffer start index
                    this.bufferStartIndex = this.parser.bufferIndex - 1;
                }

                // the overall value length has increased, as well as the piece of the value within the parser's buffer
                ++this.length;
                ++this.bufferLength;

                if (delimited)
                {
                    if (!this.delimitedStartIndex.HasValue)
                    {
                        // haven't had a delimited character yet, so set the delimited start index appropriately
                        this.delimitedStartIndex = this.length - 1;
                    }

                    this.delimitedEndIndex = this.length;
                }
            }

            // tell the value builder that the previous character parsed should not be included in the value
            // this might happen, for example, with delimiters in the value
            public void NotifyPreviousCharExcluded()
            {
                // since the value includes at least one extraneous character, we can't simply grab it straight out of the parser's buffer
                // therefore, we copy what we have demarcated so far into our local buffer and use that instead
                // TODO: that this results in an unnecessary performance hit for a fairly common scenario where data is delimited unnecessarily such as: "foo","bar"
                // in this scenario, the closing delimiter will result in this method being called and the data being copied from the parser's buffer to our local buffer, even though
                // it is likely contiguous within the parser's buffer. Have tried a couple of things to get around this, but it's resulted in awful code that I'd rather not have
                this.CopyBufferDemarcationToLocalBuffer();
            }

            // tell the value builder that the parser's buffer is about to be refilled because it is exhausted
            public void NotifyBufferRefilling()
            {
                // the value spans more than one parser buffer, so we have to save what we have demarcated so far into our local buffer and use that instead
                // of trying to just use the parser's buffer
                this.CopyBufferDemarcationToLocalBuffer();
                this.bufferStartIndex = 0;
            }

            // gets the value built by the value builder, then clears the builder ready to build the next value
            public string GetValueAndClear()
            {
                if (this.localBufferLength == 0)
                {
                    // fast path: the value fit entirely and contiguously in the parser's buffer, so we didn't need to copy anything to our local buffer
                    var buffer = this.parser.buffer;
                    var startIndex = this.delimitedStartIndex.GetValueOrDefault(this.bufferStartIndex);
                    var endIndex = startIndex + this.delimitedEndIndex.GetValueOrDefault(this.bufferLength);

                    if (!this.parser.preserveLeadingWhiteSpace)
                    {
                        while (startIndex < endIndex && IsWhiteSpace(buffer[startIndex]))
                        {
                            ++startIndex;
                        }
                    }

                    if (!this.parser.preserveTrailingWhiteSpace)
                    {
                        while (endIndex > startIndex && IsWhiteSpace(buffer[endIndex - 1]))
                        {
                            --endIndex;
                        }
                    }

                    // clear, which is slightly faster in the common case because we know localBufferLength is already zero
                    this.length = 0;
                    this.bufferLength = 0;
                    this.delimitedStartIndex = null;
                    this.delimitedEndIndex = null;

                    return new string(buffer, startIndex, endIndex - startIndex);
                }
                else
                {
                    // slow path: we had to use our local buffer to construct the value

                    // copy any outstanding data to our local buffer
                    this.CopyBufferDemarcationToLocalBuffer();

                    var buffer = this.localBuffer;
                    var startIndex = 0;
                    var endIndex = this.localBufferLength;

                    if (!this.parser.preserveLeadingWhiteSpace)
                    {
                        var stripWhiteSpaceUpToIndex = this.delimitedStartIndex.GetValueOrDefault(endIndex);

                        while (startIndex < stripWhiteSpaceUpToIndex && IsWhiteSpace(buffer[startIndex]))
                        {
                            ++startIndex;
                        }
                    }

                    if (!this.parser.preserveTrailingWhiteSpace)
                    {
                        var stripWhiteSpaceDownToIndex = this.delimitedEndIndex.GetValueOrDefault(startIndex);

                        while (endIndex > stripWhiteSpaceDownToIndex && IsWhiteSpace(buffer[endIndex - 1]))
                        {
                            --endIndex;
                        }
                    }

                    // clear
                    this.length = 0;
                    this.bufferLength = 0;
                    this.localBufferLength = 0;
                    this.delimitedStartIndex = null;
                    this.delimitedEndIndex = null;

                    return new string(buffer, startIndex, endIndex - startIndex);
                }
            }

            private void CopyBufferDemarcationToLocalBuffer()
            {
                if (this.bufferLength > 0)
                {
                    this.EnsureLocalBufferHasSufficientCapacity(this.bufferLength);

                    // copy what we demarcated in the parser's buffer into our local buffer
                    Array.Copy(this.parser.buffer, this.bufferStartIndex, this.localBuffer, this.localBufferLength, this.bufferLength);
                    this.localBufferLength += this.bufferLength;

                    // reset the demarcation of the parser's buffer back to nothing
                    this.bufferLength = 0;
                }
            }

            private void EnsureLocalBufferHasSufficientCapacity(int extraCapacityRequired)
            {
                Debug.Assert(this.localBuffer.Length >= BufferSize, "Local buffer is smaller than parser buffer, which is not supported by this method.");

                if ((this.localBufferLength + extraCapacityRequired) > this.localBuffer.Length)
                {
                    // need to allocate larger buffer
                    var newBuffer = new char[this.localBuffer.Length * 2];
                    Array.Copy(this.localBuffer, 0, newBuffer, 0, this.localBuffer.Length);
                    this.localBuffer = newBuffer;
                }
            }
        }
    }
}
