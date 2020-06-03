using System.Collections.Generic;

namespace ExcelAbstractFactory.Models
{
    public class ExcelByteModel<T>
    {
        public IList<T> ExcelItemList { get; set; }
        public string SheetName { get; set; }
        public string ColumnFixes { get; set; }
        public string ExcludeFields { get; set; }
        public string IncludeFields { get; set; }
    }
}
