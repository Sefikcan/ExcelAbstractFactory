namespace ExcelAbstractFactory
{
    public class NPOIFactory : IExcelFactory
    {
        public IExcelOperation CreateExcelBytes()
        {
            return new NPOIExcelOperation();
        }
    }
}
