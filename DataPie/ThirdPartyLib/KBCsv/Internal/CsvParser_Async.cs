namespace Kent.Boogaart.KBCsv.Internal
{
    using System.Diagnostics;
    using System.Threading.Tasks;

    // async equivalents to relevant methods in the parser
    // NOTE: changes should be made to the synchronous variants first, then ported here
    internal sealed partial class CsvParser
    {
        public async Task<int> SkipRecordsAsync(int skip)
        {
            // Performance Notes:
            //   - Checking !HasMoreRecords and exiting early degrades performance
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
                                    if (this.IsBufferEmpty && !(await this.FillBufferWithoutNotifyAsync().ConfigureAwait(false)))
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
                                if (this.IsBufferEmpty && !(await this.FillBufferWithoutNotifyAsync().ConfigureAwait(false)))
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
                        else if (!(await this.FillBufferWithoutNotifyAsync().ConfigureAwait(false)))
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

        public async Task<int> ParseRecordsAsync(HeaderRecord headerRecord, DataRecord[] buffer, int offset, int count)
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
                                if (this.IsBufferEmpty && !(await this.FillBufferAsync().ConfigureAwait(false)))
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
                            if (this.IsBufferEmpty && !(await this.FillBufferAsync().ConfigureAwait(false)))
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
                                this.activeSpecialCharacterMask = this.specialCharacterMask;
                                delimited = false;
                            }
                        }
                        else
                        {
                            // it wasn't a special character after all, so just append it
                            this.valueBuilder.NotifyPreviousCharIncluded(true);
                        }
                    }
                    else if (!(await this.FillBufferAsync().ConfigureAwait(false)))
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

        // fill the character buffer with data from the text reader
        private async Task<bool> FillBufferAsync()
        {
            Debug.Assert(this.IsBufferEmpty, "Buffer not empty.");

            this.valueBuilder.NotifyBufferRefilling();
            this.bufferEndIndex = await this.reader.ReadAsync(this.buffer, 0, BufferSize).ConfigureAwait(false);
            this.bufferIndex = 0;

            return this.bufferEndIndex > 0;
        }

        // fill the character buffer with data from the text reader. Does not notify the value builder that the fill is taking place, which is useful when the value builder is irrelevant (such as when skipping records)
        private async Task<bool> FillBufferWithoutNotifyAsync()
        {
            Debug.Assert(this.IsBufferEmpty, "Buffer not empty.");

            this.bufferEndIndex = await this.reader.ReadAsync(this.buffer, 0, BufferSize).ConfigureAwait(false);
            this.bufferIndex = 0;

            return this.bufferEndIndex > 0;
        }
    }
}