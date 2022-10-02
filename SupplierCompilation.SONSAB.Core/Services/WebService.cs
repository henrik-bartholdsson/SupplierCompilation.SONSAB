using Microsoft.Office.Interop.Excel;
using RestSharp;
using SupplierCompilation.SONSAB.Core.Dtos;
using System.Net;
using System.Xml;

namespace SupplierCompilation.SONSAB.Core.Services
{
    public class WebService : IWebService
    {
        RestClient _restClient;

        public WebService()
        {
            _restClient = new RestClient("http://ec.europa.eu/taxation_customs/vies/services/checkVatService/");
        }

        public async Task<CompanyInfoResponseDto> SendRequest(string contryCode, string VatNumber)
        {
            var request = GetRequest();

            var body = GetRequestBody(VatNumber, contryCode);

            request.AddBody(body);

            var response = await _restClient.ExecuteAsync(request);
            var xmlDoc = new XmlDocument();

            xmlDoc.LoadXml(response.Content);

            var returnObject = new CompanyInfoResponseDto { IsValid = "false" };

            if (xmlDoc.GetElementsByTagName("ns2:name").Count > 0)
            {
                returnObject.Name = xmlDoc.GetElementsByTagName("ns2:name")[0].InnerText;
            }

            if (xmlDoc.GetElementsByTagName("ns2:address").Count > 0)
            {
                returnObject.Address = xmlDoc.GetElementsByTagName("ns2:address")[0].InnerText;
            }

            if (xmlDoc.GetElementsByTagName("ns2:vatNumber").Count > 0)
            {
                returnObject.VatNumber = xmlDoc.GetElementsByTagName("ns2:vatNumber")[0].InnerText;
            }

            if (xmlDoc.GetElementsByTagName("ns2:valid").Count > 0)
            {
                returnObject.IsValid = xmlDoc.GetElementsByTagName("ns2:valid")[0].InnerText;
            }

            if (xmlDoc.GetElementsByTagName("ns2:countryCode").Count > 0)
            {
                returnObject.ContryCode = xmlDoc.GetElementsByTagName("ns2:countryCode")[0].InnerText;
            }

            return returnObject;

        }

        private string GetRequestBody(string VatNumber, string contryCode)
        {
            return @"<?xml version=""1.0"" encoding=""utf-8""?>" + "\n" +
            @"<soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" + "\n" +
            @"  <soap:Body>" + "\n" +
            @"    <checkVat xmlns=""urn:ec.europa.eu:taxud:vies:services:checkVat:types"">" + "\n" +
            $@"      <countryCode>{contryCode}</countryCode>" + "\n" +
            $@"      <vatNumber>{VatNumber}</vatNumber>" + "\n" +
            @"    </checkVat>" + "\n" +
            @"  </soap:Body>" + "\n" +
            @"</soap:Envelope>" + "\n" +
            @"";
        }

        private RestRequest GetRequest()
        {
            var request = new RestRequest("", Method.Post);
            request.AddHeader("Content-Type", "text/xml; charset=utf-8");
            request.AddHeader("Content-Length", "<calculated when request is sent>");
            request.AddHeader("User-Agent", "PostmanRuntime/7.29.2");
            request.AddHeader("Accrept", "*/*");
            request.AddHeader("Accept-Encoding", "gzip, deflate, br");
            request.AddHeader("Connection", "keep-alive");

            return request;
        }
    }
}
