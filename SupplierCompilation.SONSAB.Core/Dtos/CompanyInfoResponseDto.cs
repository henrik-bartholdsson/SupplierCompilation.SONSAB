
namespace SupplierCompilation.SONSAB.Core.Dtos
{
    public class CompanyInfoResponseDto : CompanyInfoBasisDto
    {
        public string? Name { get; set; }
        public string? Address { get; set; }
        public string? Address1 { get; set; }
        public string? Address2 { get; set; }
        public string? PostCode { get; set; }
        public string? County { get; set; }
        public string? IsValid { get; set; }

    }
}
