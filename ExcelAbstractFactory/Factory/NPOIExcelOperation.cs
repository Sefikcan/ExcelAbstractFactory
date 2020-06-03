using ExcelAbstractFactory.Models;
using System;
using System.Collections.Generic;

namespace ExcelAbstractFactory
{
    public class NPOIExcelOperation : IExcelOperation
    {
        public byte[] ToExcelBytesMultiSheets<T>(List<ExcelByteModel<T>> excelByteModel)
        {
            throw new NotImplementedException();
        }

        public byte[] ToExcelBytesOneSheet<T>(ExcelByteModel<T> excelByteModel)
        {
            throw new NotImplementedException();
        }
    }
}
