using SupplierCompilation.SONSAB.Core.Dtos;

namespace SupplierCompilation.SONSAB.Core.Services
{
    internal interface IWebService
    {
        public Task<CompanyInfoResponseDto> SendRequest(string contryCode, string VatNumber);
    }
}
