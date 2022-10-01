using SupplierCompilation.SONSAB.Core.Configurations;
using SupplierCompilation.SONSAB.Core.Dtos;

namespace SupplierCompilation.SONSAB.Core.Services
{
    public class WebService : IWebService
    {
        WebServiceConfig _config;
        public WebService(WebServiceConfig config)
        {
            _config = config;
        }

        Task<CompanyInfoResponseDto> IWebService.SendRequest(string contryCode, string VatNumber)
        {
            throw new NotImplementedException();
        }
    }
}
