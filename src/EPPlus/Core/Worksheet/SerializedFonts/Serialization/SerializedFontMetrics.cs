﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization
{
    [DebuggerDisplay("{Family} {SubFamily}")]
    internal class SerializedFontMetrics
    {
        public SerializedFontFamilies Family { get; set; }

        public FontSubFamilies SubFamily { get; set; }

        public short LineHeight { get; set; }

        public ushort UnitsPerEm { get; set; }

        public short DefaultAdvanceWidth { get; set; }

        public ushort NumberOfKerningPairs { get; set; }

        public Dictionary<char, short> AdvanceWidths { get; set; }

        public Dictionary<string, short> KerningPairs { get; set; }

        public uint GetKey()
        {
            return GetKey(Family, SubFamily);
        }

        public static uint GetKey(SerializedFontFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }

        public static uint GetKey(Font font)
        {
            var enumName = font.FontFamily.Name.Replace(" ", string.Empty);
            var values = Enum.GetValues(typeof(SerializedFontFamilies));
            var supported = false;
            foreach(var enumVal in values)
            {
                if(enumVal.ToString() == enumName)
                {
                    supported = true;
                    break;
                }
            }
            if (!supported) return uint.MaxValue;
            var family = (SerializedFontFamilies)Enum.Parse(typeof(SerializedFontFamilies), enumName);
            var subFamily = FontSubFamilies.Regular;
            switch (font.Style)
            {
                case FontStyle.Bold:
                    subFamily = FontSubFamilies.Bold;
                    break;
                case FontStyle.Italic:
                    subFamily = FontSubFamilies.Italic;
                    break;
                case FontStyle.Italic | FontStyle.Bold:
                    subFamily = FontSubFamilies.BoldItalic;
                    break;
                default:
                    break;
            }
            return GetKey(family, subFamily);
        }
    }
}
