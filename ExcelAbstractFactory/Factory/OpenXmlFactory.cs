namespace ExcelAbstractFactory
{
    public class OpenXmlFactory : IExcelFactory
    {
        public IExcelOperation CreateExcelBytes()
        {
            return new OpenXmlExcelOperation();
        }
    }
}
