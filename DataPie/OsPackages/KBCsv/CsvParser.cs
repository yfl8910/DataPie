using System;
using System.Diagnostics;
using System.IO;
using Kent.Boogaart.HelperTrinity;
using Kent.Boogaart.HelperTrinity.Extensions;

namespace Kent.Boogaart.KBCsv
{
    /// <summary>
    /// Implements the CSV parser.
    /// </summary>
    /// <remarks>
    /// This class implements the CSV parsing capabilities of KBCsv.
    /// </remarks>
    internal sealed class CsvParser : IDisposable
    {
        private static readonly ExceptionHelper _exceptionHelper = new ExceptionHelper(typeof(CsvParser));

        /// <summary>
        /// The source of the CSV data.
        /// </summary>
        private readonly TextReader _reader;

        /// <summary>
        /// See <see cref="PreserveLeadingWhiteSpace"/>.
        /// </summary>
        private bool _preserveLeadingWhiteSpace;

        /// <summary>
        /// See <see cref="PreserveTrailingWhiteSpace"/>.
        /// </summary>
        private bool _preserveTrailingWhiteSpace;

        /// <summary>
        /// See <see cref="ValueSeparator"/>.
        /// </summary>
        private char _valueSeparator;

        /// <summary>
        /// See <see cref="ValueDelimiter"/>.
        /// </summary>
        private char _valueDelimiter;

        /// <summary>
        /// Buffers CSV data.
        /// </summary>
        private readonly char[] _buffer;

        /// <summary>
        /// The current index into <see cref="_buffer"/>.
        /// </summary>
        private int _bufferIndex;

        /// <summary>
        /// The last valid index into <see cref="_buffer"/>.
        /// </summary>
        private int _bufferEndIndex;

        /// <summary>
        /// The list of values currently parsed by the parser.
        /// </summary>
        private string[] _valueList;

        /// <summary>
        /// The last valid index into <see cref="_valueList"/>.
        /// </summary>
        public int _valuesListEndIndex;

        /// <summary>
        /// The buffer of characters containing the current value.
        /// </summary>
        private char[] _valueBuffer;

        /// <summary>
        /// The last valid index into <see cref="_valueBuffer"/>.
        /// </summary>
        private int _valueBufferEndIndex;

        /// <summary>
        /// An index into <see cref="_valueBuffer"/> indicating the first character that might be removed if it is leading white-space.
        /// </summary>
        private int _valueBufferFirstEligibleLeadingWhiteSpace;

        /// <summary>
        /// An index into <see cref="_valueBuffer"/> indicating the first character that might be removed if it is trailing white-space.
        /// </summary>
        private int _valueBufferFirstEligibleTrailingWhiteSpace;

        /// <summary>
        /// <see langword="true"/> if the current value is delimited and the parser is in the delimited area.
        /// </summary>
        private bool _inDelimitedArea;

        /// <summary>
        /// The starting index of the current value part.
        /// </summary>
        private int _valuePartStartIndex;

        /// <summary>
        /// Set to <see langword="true"/> once the first record is passed (or the <see cref="CsvReader"/> decides that the first record has been passed.
        /// </summary>
        private bool _passedFirstRecord;

        /// <summary>
        /// Used to quickly recognise whether a character is potentially special or not.
        /// </summary>
        private int _specialCharacterMask;

        /// <summary>
        /// The space character.
        /// </summary>
        private const char SPACE = ' ';

        /// <summary>
        /// The tab character.
        /// </summary>
        private const char TAB = '\t';

        /// <summary>
        /// The carriage return character. Escape code is <c>\r</c>.
        /// </summary>
        private const char CR = (char) 0x0d;

        /// <summary>
        /// The line-feed character. Escape code is <c>\n</c>.
        /// </summary>
        private const char LF = (char) 0x0a;

        /// <summary>
        /// One char less than the size of the internal buffer. The extra char is used to support a faster peek operation.
        /// </summary>
        private const int BUFFER_SIZE = 2047;

        /// <summary>
        /// The default value separator.
        /// </summary>
        public const char DefaultValueSeparator = ',';

        /// <summary>
        /// The default value delimiter.
        /// </summary>
        public const char DefaultValueDelimiter = '"';

        /// <summary>
        /// Gets or sets a value indicating whether leading whitespace is to be preserved.
        /// </summary>
        public bool PreserveLeadingWhiteSpace
        {
            get
            {
                return _preserveLeadingWhiteSpace;
            }
            set
            {
                _preserveLeadingWhiteSpace = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether trailing whitespace is to be preserved.
        /// </summary>
        public bool PreserveTrailingWhiteSpace
        {
            get
            {
                return _preserveTrailingWhiteSpace;
            }
            set
            {
                _preserveTrailingWhiteSpace = value;
            }
        }

        /// <summary>
        /// Gets or sets the character that separates values in the CSV data.
        /// </summary>
        public char ValueSeparator
        {
            get
            {
                return _valueSeparator;
            }
            set
            {
                _exceptionHelper.ResolveAndThrowIf(value == _valueDelimiter, "value-separator-same-as-value-delimiter");
                _exceptionHelper.ResolveAndThrowIf(value == SPACE, "value-separator-or-value-delimiter-space");

                _valueSeparator = value;
                UpdateSpecialCharacterMask();
            }
        }

        /// <summary>
        /// Gets or sets the character that optionally delimits values in the CSV data.
        /// </summary>
        public char ValueDelimiter
        {
            get
            {
                return _valueDelimiter;
            }
            set
            {
                _exceptionHelper.ResolveAndThrowIf(value == _valueSeparator, "value-separator-same-as-value-delimiter");
                _exceptionHelper.ResolveAndThrowIf(value == SPACE, "value-separator-or-value-delimiter-space");

                _valueDelimiter = value;
                UpdateSpecialCharacterMask();
            }
        }

        /// <summary>
        /// Gets a value indicating whether the parser's buffer contains more records in addition to those already parsed.
        /// </summary>
        public bool HasMoreRecords
        {
            get
            {
                if (_bufferIndex < _bufferEndIndex)
                {
                    //the buffer isn't empty so there must be more records
                    return true;
                }
                else
                {
                    //the buffer is empty so peek into the reader to see whether there is more data
                    return (_reader.Peek() != -1);
                }			
            }
        }

        /// <summary>
        /// Gets a value indicating whether the parser has passed the first record in the input source.
        /// </summary>
        public bool PassedFirstRecord
        {
            get
            {
                return _passedFirstRecord;
            }
            set
            {
                _passedFirstRecord = value;
            }
        }
        
        /// <summary>
        /// Constructs and inititialises an instance of <c>CsvParser</c> with the details provided.
        /// </summary>
        /// <param name="reader">
        /// The instance of <see cref="TextReader"/> from which CSV data will be read.
        /// </param>
        public CsvParser(TextReader reader)
        {
            reader.AssertNotNull("reader");

            _reader = reader;
            //the extra char is used to facilitate a faster peek operation
            _buffer = new char[BUFFER_SIZE + 1];
            _valueList = new string[16];
            _valueBuffer = new char[128];
            _valuePartStartIndex = -1;
            //set defaults
            _valueSeparator = DefaultValueSeparator;
            _valueDelimiter = DefaultValueDelimiter;
            //get the default special character mask
            UpdateSpecialCharacterMask();
        }

        /// <summary>
        /// Efficiently skips the next CSV record.
        /// </summary>
        /// <returns>
        /// <see langword="true"/> if a record was successfully skipped, otherwise <see langword="false"/>.
        /// </returns>
        public bool SkipRecord()
        {
            if (HasMoreRecords)
            {
                //taking a local copy allows optimisations that otherwise could not be performed because the CLR knows that no other thread
                //can touch our local reference
                char[] buffer = _buffer;

                while (true)
                {
                    if (_bufferIndex != _bufferEndIndex)
                    {
                        char c = buffer[_bufferIndex++];

                        if ((c & _specialCharacterMask) == c)
                        {
                            if (!_inDelimitedArea)
                            {
                                if (c == _valueDelimiter)
                                {
                                    //found a delimiter so enter delimited area and set the start index for the value
                                    _inDelimitedArea = true;
                                }
                                else if (c == CR)
                                {
                                    if (buffer[_bufferIndex] == LF)
                                    {
                                        SwallowChar();
                                    }

                                    return true;
                                }
                                else if (c == LF)
                                {
                                    return true;
                                }
                            }
                            else if (c == _valueDelimiter)
                            {
                                if (buffer[_bufferIndex] == _valueDelimiter)
                                {
                                    // delimiter is escaped, so swallow the escape char
                                    SwallowChar();
                                }
                                else
                                {
                                    //delimiter isn't escaped so we are no longer in a delimited area
                                    _inDelimitedArea = false;
                                }
                            }
                        }
                    }
                    else if (!FillBufferIgnoreValues())
                    {
                        //data exhausted - get out of here
                        return true;
                    }
                }
            }
            else
            {
                //no more records - can't skip
                return false;
            }
        }

        /// <summary>
        /// Reads and parses the CSV into a <c>string</c> array containing the values contained in a single CSV record.
        /// </summary>
        /// <returns>
        /// An array of field values for the record, or <see langword="null"/> if no record was found.
        /// </returns>
        public string[] ParseRecord()
        {
            _valuesListEndIndex = 0;
            char c = char.MinValue;
            //taking a local copy allows optimisations that otherwise could not be performed because the CLR knows that no other thread
            //can touch our local reference
            char[] buffer = _buffer;

            while (true)
            {
                if (_bufferIndex != _bufferEndIndex)
                {
                    if (_valuePartStartIndex == -1)
                    {
                        _valuePartStartIndex = _bufferIndex;
                    }

                    c = buffer[_bufferIndex++];

                    if ((c & _specialCharacterMask) == c)
                    {
                        if (!_inDelimitedArea)
                        {
                            if (c == _valueDelimiter)
                            {
                                //found a delimiter so enter delimited area and set the start index for the value
                                _inDelimitedArea = true;
                                CloseValuePartExcludeCurrent();
                                _valuePartStartIndex = _bufferIndex;

                                //we have to make sure that delimited text isn't stripped if it is white-space
                                if (_valueBufferFirstEligibleLeadingWhiteSpace == 0)
                                {
                                    _valueBufferFirstEligibleLeadingWhiteSpace = _valueBufferEndIndex;
                                }
                            }
                            else if (c == _valueSeparator)
                            {
                                CloseValue(false);
                            }
                            else if (c == CR)
                            {
                                CloseValue(false);

                                if (buffer[_bufferIndex] == LF)
                                {
                                    SwallowChar();
                                }

                                break;
                            }
                            else if (c == LF)
                            {
                                CloseValue(false);
                                break;
                            }
                        }
                        else if (c == _valueDelimiter)
                        {
                            if (buffer[_bufferIndex] == _valueDelimiter)
                            {
                                CloseValuePart();
                                SwallowChar();
                                _valuePartStartIndex = _bufferIndex;
                            }
                            else
                            {
                                //delimiter isn't escaped so we are no longer in a delimited area
                                _inDelimitedArea = false;
                                CloseValuePartExcludeCurrent();
                                _valuePartStartIndex = _bufferIndex;
                                //we have to make sure that delimited text isn't stripped if it is white-space
                                _valueBufferFirstEligibleTrailingWhiteSpace = _valueBufferEndIndex;
                            }
                        }
                    }
                }
                else if (!FillBuffer())
                {
                    //special case: if the last character was a separator we need to add a blank value. eg. CSV "Value," will result in "Value", ""
                    if (c == _valueSeparator)
                    {
                        AddValue(string.Empty);
                    }

                    //data exhausted - get out of loop
                    break;
                }
            }

            //this will return null if there are no values
            return GetValues();
        }

        /// <summary>
        /// Closes the current value part.
        /// </summary>
        private void CloseValuePart()
        {
            AppendToValue(_valuePartStartIndex, _bufferIndex);
            _valuePartStartIndex = -1;
        }

        /// <summary>
        /// Closes the current value part, but excludes the current character from the value part.
        /// </summary>
        private void CloseValuePartExcludeCurrent()
        {
            AppendToValue(_valuePartStartIndex, _bufferIndex - 1);
        }

        /// <summary>
        /// Closes the current value by adding it to the list of values in the current record. Assumes that there is actually a value to add, either in <c>_value</c> or in
        /// <see cref="_buffer"/> starting at <see cref="_valuePartStartIndex"/> and ending at <see cref="_bufferIndex"/>.
        /// </summary>
        /// <param name="includeCurrentChar">
        /// If <see langword="true"/>, the current character is included in the value. Otherwise, it is excluded.
        /// </param>
        private void CloseValue(bool includeCurrentChar)
        {
            int endIndex = _bufferIndex;

            if ((!includeCurrentChar) && (endIndex > _valuePartStartIndex))
            {
                endIndex -= 1;
            }

            Debug.Assert(_valuePartStartIndex >= 0, "_valuePartStartIndex must be > 0");
            Debug.Assert(_valuePartStartIndex <= _bufferIndex, "_valuePartStartIndex must be less than or equal to _bufferIndex (" + _valuePartStartIndex + " > " + _bufferIndex + ")");
            Debug.Assert(_valuePartStartIndex <= endIndex, "_valuePartStartIndex must be less than or equal to endIndex (" + _valuePartStartIndex + " > " + endIndex + ")");

            if (_valueBufferEndIndex == 0)
            {
                if (endIndex == 0)
                {
                    AddValue(string.Empty);
                }
                else
                {
                    //the value did not require the use of the ValueBuilder
                    int startIndex = _valuePartStartIndex;
                    //taking a local copy allows optimisations that otherwise could not be performed because the CLR knows that no other thread
                    //can touch our local reference
                    char[] buffer = _buffer;

                    if (!_preserveLeadingWhiteSpace)
                    {
                        //strip all leading white-space
                        while ((startIndex < endIndex) && (IsWhiteSpace(buffer[startIndex])))
                        {
                            ++startIndex;
                        }
                    }

                    if (!_preserveTrailingWhiteSpace)
                    {
                        //strip all trailing white-space
                        while ((endIndex > startIndex) && (IsWhiteSpace(buffer[endIndex - 1])))
                        {
                            --endIndex;
                        }
                    }

                    AddValue(new string(buffer, startIndex, endIndex - startIndex));
                }
            }
            else
            {
                //we needed the ValueBuilder to compose the value
                AppendToValue(_valuePartStartIndex, endIndex);

                if (!_preserveLeadingWhiteSpace || !_preserveTrailingWhiteSpace)
                {
                    //strip all white-space prior to _valueBufferFirstEligibleLeadingWhiteSpace and after _valueBufferFirstEligibleTrailingWhiteSpace
                    AddValue(GetValue(_valueBufferFirstEligibleLeadingWhiteSpace, _valueBufferFirstEligibleTrailingWhiteSpace));
                }
                else
                {
                    AddValue(GetValue());
                }

                _valueBufferEndIndex = 0;
                _valueBufferFirstEligibleLeadingWhiteSpace = 0;
            }

            _valuePartStartIndex = -1;
        }

        /// <summary>
        /// Determines whether <paramref name="c"/> is white-space.
        /// </summary>
        /// <param name="c">
        /// The character to check.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="c"/> is white-space, otherwise <see langword="false"/>.
        /// </returns>
        internal static bool IsWhiteSpace(char c)
        {
            return ((c == SPACE) || (c == TAB));
        }

        /// <summary>
        /// Fills that data buffer. Assumes that the buffer is already exhausted.
        /// </summary>
        /// <returns>
        /// <see langword="true"/> if data was read into the buffer, otherwise <see langword="false"/>.
        /// </returns>
        private bool FillBuffer()
        {
            Debug.Assert(_bufferIndex == _bufferEndIndex);

            //may need to close a value or value part depending on the state of the stream
            if (_valuePartStartIndex != -1)
            {
                if (_reader.Peek() != -1)
                {
                    CloseValuePart();
                    _valuePartStartIndex = 0;
                }
                else
                {
                    CloseValue(true);
                }
            }

            _bufferEndIndex = _reader.Read(_buffer, 0, BUFFER_SIZE);
            //this is possible because the buffer is one char bigger than BUFFER_SIZE. This fact is used to implement a faster peek operation
            _buffer[_bufferEndIndex] = (char) _reader.Peek();
            _bufferIndex = 0;
            _passedFirstRecord = true;
            return (_bufferEndIndex > 0);
        }

        /// <summary>
        /// Fills the buffer with data, but does not bother with closing values. This is used from the <see cref="SkipRecord"/> method,
        /// since that does not concern itself with values.
        /// </summary>
        /// <returns>
        /// <see langword="true"/> if data was read into the buffer, otherwise <see langword="false"/>.
        /// </returns>
        private bool FillBufferIgnoreValues()
        {
            Debug.Assert(_bufferIndex == _bufferEndIndex);
            _bufferEndIndex = _reader.Read(_buffer, 0, BUFFER_SIZE);
            //this is possible because the buffer is one char bigger than BUFFER_SIZE. This fact is used to implement a faster peek operation
            _buffer[_bufferEndIndex] = (char) _reader.Peek();
            _bufferIndex = 0;
            _passedFirstRecord = true;
            return (_bufferEndIndex > 0);
        }


        /// <summary>
        /// Swallows the current character in the data buffer. Assumes that there is a character to swallow, but refills the buffer if necessary.
        /// </summary>
        private void SwallowChar()
        {
            if (_bufferIndex < BUFFER_SIZE)
            {
                //in this case there are still unread chars in the buffer so just skip one
                ++_bufferIndex;
            }
            else if (_bufferIndex < _bufferEndIndex)
            {
                //in this case we are pointing to the second-to-last char in the buffer, so we need to refill the buffer since the last char is a peeked char
                FillBuffer();
            }
            else
            {
                //in this case we are pointing to the last char in the buffer, which is a peeked char. therefore, we need to refill and skip past that char
                FillBuffer();
                ++_bufferIndex;
            }
        }

        /// <summary>
        /// Disposes of this <c>CsvParser</c> instance.
        /// </summary>
        void IDisposable.Dispose()
        {
            Close();
        }

        /// <summary>
        /// Closes this <c>CsvParser</c> instance and releases all resources acquired by it.
        /// </summary>
        public void Close()
        {
            if (_reader != null)
            {
                _reader.Close();
            }
        }
        
        /// <summary>
        /// Adds a value to the value list.
        /// </summary>
        /// <param name="val">
        /// The value to add.
        /// </param>
        private void AddValue(string val)
        {
            EnsureValueListCapacity();
            _valueList[_valuesListEndIndex++] = val;
        }

        /// <summary>
        /// Gets an array of values that have been added to <see cref="_valueList"/>.
        /// </summary>
        /// <returns>
        /// An array of type <c>string</c> containing all the values in the value list, or <see langword="null"/> if there are no values in the list.
        /// </returns>
        private string[] GetValues()
        {
            if (_valuesListEndIndex > 0)
            {
                string[] retVal = new string[_valuesListEndIndex];

                for (int i = 0; i < _valuesListEndIndex; ++i)
                {
                    retVal[i] = _valueList[i];
                }

                return retVal;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Ensures the value list contains enough space for another value, and increases its size if not.
        /// </summary>
        private void EnsureValueListCapacity()
        {
            if (_valuesListEndIndex == _valueList.Length)
            {
                string[] newBuffer = new string[_valueList.Length * 2];

                for (int i = 0; i < _valuesListEndIndex; ++i)
                {
                    newBuffer[i] = _valueList[i];
                }

                _valueList = newBuffer;
            }
        }
        
        /// <summary>
        /// Appends the specified characters from <see cref="_buffer"/> onto the end of the current value.
        /// </summary>
        /// <param name="startIndex">
        /// The index at which to begin copying.
        /// </param>
        /// <param name="endIndex">
        /// The index at which to cease copying. The character at this index is not copied.
        /// </param>
        private void AppendToValue(int startIndex, int endIndex)
        {
            EnsureValueBufferCapacity(endIndex - startIndex);
            char[] valueBuffer = _valueBuffer;
            char[] buffer = _buffer;

            //profiling revealed a loop to be faster than Array.Copy
            //in addition, profiling revealed that taking a local copy of the _buffer reference impedes performance here
            for (int i = startIndex; i < endIndex; ++i)
            {
                valueBuffer[_valueBufferEndIndex++] = buffer[i];
            }
        }
        
        /// <summary>
        /// Gets the current value.
        /// </summary>
        /// <returns></returns>
        private string GetValue()
        {
            return new string(_valueBuffer, 0, _valueBufferEndIndex);
        }

        /// <summary>
        /// Gets the current value, optionally removing trailing white-space.
        /// </summary>
        /// <param name="valueBufferFirstEligibleLeadingWhiteSpace">
        /// The index of the first character that cannot possibly be leading white-space.
        /// </param>
        /// <param name="valueBufferFirstEligibleTrailingWhiteSpace">
        /// The index of the first character that may be trailing white-space.
        /// </param>
        /// <returns>
        /// An instance of <c>string</c> containing the resultant value.
        /// </returns>
        private string GetValue(int valueBufferFirstEligibleLeadingWhiteSpace, int valueBufferFirstEligibleTrailingWhiteSpace)
        {
            int startIndex = 0;
            int endIndex = _valueBufferEndIndex - 1;

            if (!_preserveLeadingWhiteSpace)
            {
                while ((startIndex < valueBufferFirstEligibleLeadingWhiteSpace) && (IsWhiteSpace(_valueBuffer[startIndex])))
                {
                    ++startIndex;
                }
            }

            if (!_preserveTrailingWhiteSpace)
            {
                while ((endIndex >= valueBufferFirstEligibleTrailingWhiteSpace) && (IsWhiteSpace(_valueBuffer[endIndex])))
                {
                    --endIndex;
                }
            }

            return new string(_valueBuffer, startIndex, endIndex - startIndex + 1);
        }

        /// <summary>
        /// Ensures the value buffer contains enough space for <paramref name="count"/> more characters.
        /// </summary>
        private void EnsureValueBufferCapacity(int count)
        {
            if ((_valueBufferEndIndex + count) > _valueBuffer.Length)
            {
                char[] newBuffer = new char[Math.Max(_valueBuffer.Length * 2, (count >> 1) << 2)];

                //profiling revealed a loop to be faster than Array.Copy, despite Array.Copy having an internal implementation
                for (int i = 0; i < _valueBufferEndIndex; ++i)
                {
                    newBuffer[i] = _valueBuffer[i];
                }

                _valueBuffer = newBuffer;
            }
        }

        /// <summary>
        /// Updates the mask used to quickly filter out non-special characters.
        /// </summary>
        private void UpdateSpecialCharacterMask()
        {
            _specialCharacterMask = _valueSeparator | _valueDelimiter | CR | LF;
        }
    }
}
