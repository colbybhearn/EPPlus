﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/30/2023         EPPlus Software AB       Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml.LoadFunctions.Params;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromFixedWidthText : LoadFromTextBase
    {
        protected ExcelTextFormatFixedWidth _format;

        public LoadFromFixedWidthText(ExcelRangeBase range, string text, ExcelTextFormatFixedWidth Format) 
            : base(range, text)
        {
            _format = Format;
        }
        public override ExcelRangeBase Load()
        {
            if (string.IsNullOrEmpty(_text))
            {
                var r = _worksheet.Cells[_range._fromRow, _range._fromCol];
                r.Value = "";
                return r;
            }

            string[] lines;
            lines = SplitLines(_text, _format.EOL);
            var col = 0;
            var maxCol = col;
            var row = 0;
            var lineNo = 1;
            foreach (string line in lines)
            {
                if (lineNo > _format.SkipLinesBeginning && lineNo <= lines.Length - _format.SkipLinesEnd)
                {
                    if (string.IsNullOrEmpty(line))
                    {
                        continue;
                    }
                    var items = new List<object>();
                    var isText = false;
                    int readLength = 0;
                    col = 0;
                    for (int i = 0; i < _format.ColumnLengths.Length; i++)
                    {
                        string content;
                        if (i == 0)
                        {
                            content = line.Substring(0, _format.ColumnLengths[i]);
                            readLength += _format.ColumnLengths[i];
                        }
                        else
                        {
                            var v = line.Length;
                            if (readLength + _format.ColumnLengths[i] >= v)
                            {
                                content = line.Substring(readLength + 1);
                            }
                            else
                            {
                                content = line.Substring(readLength, _format.ColumnLengths[i]);
                                readLength += _format.ColumnLengths[i];
                            }
                        }
                        content = content.Trim();
                        if (_format.UseColumns == null || (_format.UseColumns != null && _format.UseColumns[i]))
                        {
                            items.Add(ConvertData(_format, content.Trim(), col, isText));
                        }
                        col++;
                    }
                    _worksheet._values.SetValueRow_Value(_range._fromRow + row, _range._fromCol, items);
                    if (col > maxCol) maxCol = col;
                    row++;
                }
                lineNo++;
            }
            return _worksheet.Cells[_range._fromRow, _range._fromCol, _range._fromRow + row - 1, _range._fromCol + maxCol-1];
        }

    }
}
