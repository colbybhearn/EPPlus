﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/15/2024         EPPlus Software AB       Initial release EPPlus 7.2
 *************************************************************************************************/
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Metadata;
using OfficeOpenXml.FormulaParsing.FormulaExpressions;
using OfficeOpenXml.FormulaParsing.Ranges;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
    [FunctionMetadata(
        Category = ExcelFunctionCategory.Text,
        EPPlusVersion = "7.2",
        Description = "Get the text before delimiter",
        SupportsArrays = false)]
    internal class TextBefore : ExcelFunctionTextBase
    {
        public override int ArgumentMinLength => 2;
        public override string NamespacePrefix => "_xlfn.";

        public override CompileResult Execute(IList<FunctionArgument> arguments, ParsingContext context)
        {
            var text = ArgToString(arguments, 0);
            var delimiters = ArgDelimiterCollectionToString(arguments, 1, out CompileResult error);
            if (error != null) return error;
            var instanceNum = 1;
            var matchMode = "0";
            var matchEnd = 0;
            var ifNotFound = "#N/A";
            var resultString = string.Empty;
            if(arguments.Count > 2)
            {
                instanceNum = ArgToInt(arguments, 2, RoundingMethod.Convert);
                if(instanceNum == 0)
                {
                    instanceNum = 1;
                }
            }
            if (arguments.Count > 3)
            {
                matchMode = ArgToString(arguments, 3);
                if (matchMode == "1")
                {
                    delimiters += delimiters.ToLower() + delimiters.ToUpper();
                }
            }
            if (arguments.Count > 4)
            {
                matchEnd = ArgToInt(arguments, 4, RoundingMethod.Convert);
            }
            if (arguments.Count > 5)
            {
                ifNotFound = ArgToString(arguments, 5);
            }


            int length = 0;
            int instances = 0;
            if (instanceNum < 0)
            {
                for (int i = text.Length - 1; i >= 0; i--)
                {
                    char c = text[i];
                    if (delimiters.Contains(c))
                    {
                        instances--;
                        length = i;
                        if (instances == instanceNum) break;
                    }
                }

                if (instances > instanceNum && matchEnd == 0)
                {
                    if (ifNotFound != "#N/A")
                        return CreateResult(ifNotFound, DataType.String);
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
                if (matchEnd == 1 && instances - instanceNum == 1)
                {
                    return CreateResult(text, DataType.String);
                }
                else if (matchEnd == 1 && instances - instanceNum > 1)
                {
                    if (ifNotFound != "#N/A")
                        return CreateResult(ifNotFound, DataType.String);
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
            }
            else
            {
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    if (delimiters.Contains(c))
                    {
                        instances++;
                        length = i;
                        if (instances == instanceNum)
                        {
                            break;
                        }
                    }
                }
                if (instances < instanceNum && matchEnd == 0)
                {
                    if (ifNotFound != "#N/A")
                        return CreateResult(ifNotFound, DataType.String);
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
                if (matchEnd == 1 && instances - instanceNum == -1)
                {
                    return CreateResult(text, DataType.String);
                }
                else if(matchEnd == 1 && instances - instanceNum < -1)
                {
                    if (ifNotFound != "#N/A")
                        return CreateResult(ifNotFound, DataType.String);
                    return CompileResult.GetErrorResult(eErrorType.NA);
                }
            }
            length = length--;
            if(length <= 0)
            {
                if (ifNotFound != "#N/A")
                    return CreateResult(ifNotFound, DataType.String);
                return CompileResult.GetErrorResult(eErrorType.NA);
            }
            resultString = text.Substring(0, length);
            return CreateResult(resultString, DataType.String);
        }
    }
}
