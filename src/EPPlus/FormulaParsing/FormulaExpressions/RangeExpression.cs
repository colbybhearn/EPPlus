﻿using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.FormulaExpressions
{
    [DebuggerDisplay("RangeExpression: {_addressInfo.Address}")]
    internal class RangeExpression : Expression
    {
        protected FormulaRangeAddress _addressInfo;
        internal RangeExpression(CompileResult result, ParsingContext ctx) : base(ctx)
        {
            _cachedCompileResult = result;
            _addressInfo = result.Address;
        }
        internal RangeExpression(FormulaRangeAddress address) : base(address._context)
        {
            _addressInfo = address;
        }
        internal RangeExpression(string address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            Init(new FormulaRangeAddress(ctx, address), ctx, externalReferenceIx, worksheetIx);
        }
        internal RangeExpression(ExcelAddressBase address, ParsingContext ctx, short externalReferenceIx, int worksheetIx) : base(ctx)
        {
            Init(address.AsFormulaRangeAddress(ctx), ctx, externalReferenceIx, worksheetIx);
        }

        private void Init(FormulaRangeAddress address, ParsingContext ctx, short externalReferenceIx, int worksheetIx)
        {
            _addressInfo = address;
            _addressInfo.ExternalReferenceIx = externalReferenceIx;
            _addressInfo.WorksheetIx = (worksheetIx == int.MinValue ? ctx.CurrentCell.WorksheetIx : worksheetIx);
        }

        internal override ExpressionType ExpressionType => ExpressionType.CellAddress;
        public override CompileResult Compile()
        {
            if (_cachedCompileResult == null)
            {
                if(_addressInfo.ExternalReferenceIx < 1)
                {
                    if (_addressInfo.IsSingleCell)
                    {
                        if (_addressInfo.WorksheetIx == -1)
                        {
                            _cachedCompileResult = CompileResult.GetErrorResult(eErrorType.Ref);
                        }
                        else
                        {
                            var ws = Context.Package.Workbook.GetWorksheetByIndexInList(_addressInfo.WorksheetIx);
                            var v = ws.GetValue(_addressInfo.FromRow, _addressInfo.FromCol); //Use GetValue to get richtext values.
                            _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo);
                            _cachedCompileResult.IsHiddenCell = ws.IsRowHidden(_addressInfo.FromRow);
                        }
                    }
                    else
                    {
                        _cachedCompileResult = new AddressCompileResult(new RangeInfo(_addressInfo), DataType.ExcelRange, _addressInfo);
                    }
                }
                else
                {
                    var ri = _addressInfo.GetAsRangeInfo();
                    if (ri.GetNCells() > 1)
                    {
                        _cachedCompileResult = new AddressCompileResult(ri, DataType.ExcelRange, _addressInfo);
                    }
                    else
                    {
                        var v = ri.GetOffset(0, 0);
                        _cachedCompileResult = CompileResultFactory.Create(v, _addressInfo);
                    }
                }
            }
            return _cachedCompileResult;
        }

        public override Expression Negate()
        {
            if (_cachedCompileResult == null)
            {
                Compile();
            }
            return new RangeExpression(_cachedCompileResult.Negate(), Context);
        }
        internal override ExpressionStatus Status
        {
            get;
            set;
        } = ExpressionStatus.IsAddress;
        internal override Expression CloneWithOffset(int row, int col)
        {
            var ai = new FormulaRangeAddress(Context)
            {
                ExternalReferenceIx = _addressInfo.ExternalReferenceIx,
                WorksheetIx = _addressInfo.WorksheetIx,
                FromRow = (_addressInfo.FixedFlag & FixedFlag.FromRowFixed) == FixedFlag.FromRowFixed ? _addressInfo.FromRow : _addressInfo.FromRow + row,
                ToRow = (_addressInfo.FixedFlag & FixedFlag.ToRowFixed) == FixedFlag.ToRowFixed ? _addressInfo.ToRow : _addressInfo.ToRow + row,
                FromCol = (_addressInfo.FixedFlag & FixedFlag.FromColFixed) == FixedFlag.FromColFixed ? _addressInfo.FromCol : _addressInfo.FromCol + col,
                ToCol = (_addressInfo.FixedFlag & FixedFlag.ToColFixed) == FixedFlag.ToColFixed ? _addressInfo.ToCol : _addressInfo.ToCol + col,
            };
            return new RangeExpression(ai)
            {
                Status = Status,                
                Operator= Operator
            };
        }
        public override FormulaRangeAddress[] GetAddress() 
        {
            return [_addressInfo.Clone()];
        }
        internal override void MergeAddress(string address)
        {
            ExcelCellBase.GetRowColFromAddress(address, out int fromRow, out int fromCol, out int toRow, out int toCol, out bool fixedFromRow, out bool fixedFromCol, out bool fixedToRow, out bool fixedToCol);

            if (_addressInfo.FromRow > fromRow)
            {
                _addressInfo.FromRow = fromRow;
                SetFixedFlag(fixedFromRow, FixedFlag.FromRowFixed);
            }
            if (_addressInfo.ToRow < toRow)
            {
                _addressInfo.ToRow = toRow;
                SetFixedFlag(fixedToRow, FixedFlag.ToRowFixed);
            }
            if (_addressInfo.FromCol > fromCol)
            {
                _addressInfo.FromCol = fromCol;
                SetFixedFlag(fixedFromCol, FixedFlag.FromColFixed);
            }
            if (_addressInfo.ToCol < toCol)
            {
                _addressInfo.ToCol = toCol;
                SetFixedFlag(fixedToCol, FixedFlag.ToColFixed);
            }
        }

        private void SetFixedFlag(bool setFlag, FixedFlag flag)
        {
            if (setFlag)
            {
                _addressInfo.FixedFlag |= flag;
            }
            else
            {
                _addressInfo.FixedFlag &= ~flag;
            }
        }
    }
}
