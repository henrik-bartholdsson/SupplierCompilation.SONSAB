﻿
namespace SupplierCompilation.SONSAB.Core.Dtos
{
    internal class CompanyInfoResponseDto : CompanyInfoBasisDto
    {
        public string Name { get; set; }
        public string Address { get; set; }
        public string IsValid { get; set; }

    }
}