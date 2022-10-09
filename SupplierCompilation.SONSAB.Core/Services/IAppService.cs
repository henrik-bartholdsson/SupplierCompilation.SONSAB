namespace SupplierCompilation.SONSAB.Core.Services
{
    public interface IAppService
    {
        public void ProcessVatFile(string filePath);
        public void SetColumn(string column);

    }
}
