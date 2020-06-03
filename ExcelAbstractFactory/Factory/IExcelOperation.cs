using ExcelAbstractFactory.Models;
using System.Collections.Generic;

namespace ExcelAbstractFactory
{
    public interface IExcelOperation
    {
        //Tek sheet ile excel data'sı oluşturur
        byte[] ToExcelBytesOneSheet<T>(ExcelByteModel<T> excelByteModel);

        //Çoklu sheet ile excel data oluşturur.
        byte[] ToExcelBytesMultiSheets<T>(List<ExcelByteModel<T>> excelByteModel);
    }
}
