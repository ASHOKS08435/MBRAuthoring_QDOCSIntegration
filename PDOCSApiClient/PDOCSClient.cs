using Newtonsoft.Json;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace PDOCSApiClient
{
    public class APIResponse
    {
        public bool Success { get; set; }
        public string Response { get; set; }
        public string ErrorMessage { get; set; }
        public string ErrorDetails { get; set; }
        public HttpStatusCode StatusCode { get; set; }
    }

    public class ErrorResponse
    {
        [JsonProperty("Error_Description")]
        public string ErrorDescription { get; set; }
        public string Error { get; set; }

        [JsonProperty("Error Details")]
        public string ErrorDetails { get; set; }
    }

    public class ResponseData
    {
        public Row[] Rows { get; set; }
        public int Offset { get; set; }
        public int TotalLength { get; set; }
    }

    public class Row
    {
        [JsonProperty("properties")]
        public RowProperties Properties { get; set; }
    }

    public class RowProperties
    {
        [JsonProperty("r_object_id")]
        public string DocObjectId { get; set; }

        [JsonProperty("title")]
        public string Title { get; set; }
    }

    public class MulesoftTokenResponse
    {
        public string access_token { get; set; }
        public string token_type { get; set; }
        public int expires_in { get; set; }
    }

    public class PDOCSClient
    {
        private readonly HttpClient _httpClient;
        private CookieContainer cookieContainer;
        private readonly string _accessTokenUrl;
        public string setCookie;
        HttpResponseMessage responseMessage = new HttpResponseMessage(HttpStatusCode.BadRequest);
        string responseBody = null;
        private string _accessToken;

        public PDOCSClient(string accessTokenUrl, string clientId, string clientSecret)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            cookieContainer = new CookieContainer();
            var handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer
            };
            _httpClient = new HttpClient(handler);
            _accessTokenUrl = accessTokenUrl;
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.ASCII.GetBytes($"{clientId}:{clientSecret}")));
        }

        public async Task<APIResponse> GetAccessTokenAsync()
        {
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Post, _accessTokenUrl)
                {
                    Content = new StringContent("grant_type=client_credentials", Encoding.UTF8, "application/x-www-form-urlencoded")
                };

                responseMessage = await _httpClient.SendAsync(request);

                responseBody = await responseMessage.Content.ReadAsStringAsync();
                if (responseMessage.IsSuccessStatusCode)
                {
                    MulesoftTokenResponse tokenResponse = JsonConvert.DeserializeObject<MulesoftTokenResponse>(responseBody);
                    _accessToken = tokenResponse.access_token;
                    return new APIResponse { Success = true, StatusCode = responseMessage.StatusCode };
                }

                ErrorResponse errResponse = JsonConvert.DeserializeObject<ErrorResponse>(responseBody);
                return new APIResponse { Success = false, StatusCode = responseMessage.StatusCode, ErrorMessage = "Mulesoft Login Failure", ErrorDetails = errResponse.ErrorDescription + " " + errResponse.ErrorDetails + " " + errResponse.ErrorDetails };
            }
            catch (Exception ex)
            {
                return new APIResponse { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.StackTrace };
            }
        }

        public async Task<APIResponse> PostLoginAsAsync(string url, string requestBody)
        {
            try
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                responseMessage = await _httpClient.PostAsync(url, content);

                responseBody = await responseMessage.Content.ReadAsStringAsync();

                if (responseMessage.IsSuccessStatusCode)
                {
                    setCookie = responseMessage.Headers.GetValues("Set-Cookie").FirstOrDefault();

                    return new APIResponse { Success = true, StatusCode = responseMessage.StatusCode };
                }

                ErrorResponse errResponse = JsonConvert.DeserializeObject<ErrorResponse>(responseBody);
                return new APIResponse { Success = false, StatusCode = responseMessage.StatusCode, ErrorMessage = "PDOCS Login Failure - User/Service account is not a valid PDOCS User", ErrorDetails = errResponse.ErrorDescription + " " + errResponse.ErrorDetails + " " + errResponse.Error };
            }

            catch (Exception ex)
            {
                return new APIResponse { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.StackTrace };
            }
        }

        public async Task<APIResponse> PostGetQueryAsync(string url, string requestBody, bool isUserValidation)
        {
            try
            {
                var uri = new Uri(url);
                foreach (var cookie in setCookie.Split(';'))
                {
                    var cookieParts = cookie.Split('=');
                    if (cookieParts.Length == 2)
                    {
                        cookieContainer.Add(uri, new Cookie(cookieParts[0].Trim(), cookieParts[1].Trim()));
                    }
                }

                var reqMsg = new HttpRequestMessage(HttpMethod.Post, url);
                reqMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
                reqMsg.Content = content;
                responseMessage = await _httpClient.SendAsync(reqMsg);

                responseBody = await responseMessage.Content.ReadAsStringAsync();

                if (responseMessage.IsSuccessStatusCode)
                {
                    ResponseData response = JsonConvert.DeserializeObject<ResponseData>(responseBody);

                    if (response.TotalLength > 0)
                    {
                        return new APIResponse { Success = true, StatusCode = responseMessage.StatusCode, Response = response.Rows[0].Properties.DocObjectId };
                    }
                    else
                    {
                        return new APIResponse { Success = false, StatusCode = responseMessage.StatusCode, ErrorMessage = isUserValidation ? "User/Service account is not a valid PDOCS User" : "Document not found" };
                    }

                }
                else
                {
                    ErrorResponse errResponse = JsonConvert.DeserializeObject<ErrorResponse>(responseBody);
                    return new APIResponse { Success = false, StatusCode = responseMessage.StatusCode, ErrorMessage = "User Validation Failure on PDOCS", ErrorDetails = errResponse.ErrorDescription + " " + errResponse.ErrorDetails + " " + errResponse.ErrorDetails };
                }
            }
            catch (Exception ex)
            {
                return new APIResponse { Success = false, ErrorMessage = ex.Message, ErrorDetails = ex.StackTrace };
            }
        }

        public async Task ExportDocAsync(string url, string filePath)
        {
            try
            {
                var uri = new Uri(url);
                foreach (var cookie in setCookie.Split(';'))
                {
                    var cookieParts = cookie.Split('=');
                    if (cookieParts.Length == 2)
                    {
                        cookieContainer.Add(uri, new Cookie(cookieParts[0].Trim(), cookieParts[1].Trim()));
                    }
                }

                var reqMsg = new HttpRequestMessage(HttpMethod.Get, url);
                reqMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _accessToken);

                var response = await _httpClient.SendAsync(reqMsg);
                response.EnsureSuccessStatusCode();

                var fileBytes = await response.Content.ReadAsByteArrayAsync();
                File.WriteAllBytes(filePath, fileBytes);
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
    //public class PDOCSClient
    //{
    //    private readonly HttpClient _httpClient;
    //    private CookieContainer cookieContainer;
    //    private string _accessTokenUrl;
    //    public string setCookie;

    //    public PDOCSClient(string accessTokenUrl, string clientId, string clientSecret)
    //    {
    //        cookieContainer = new CookieContainer();
    //        var handler = new HttpClientHandler
    //        {
    //            CookieContainer = cookieContainer
    //        };
    //        _httpClient = new HttpClient(handler);
    //        _accessTokenUrl = accessTokenUrl;
    //        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.ASCII.GetBytes($"{clientId}:{clientSecret}")));
    //    }

    //    public async Task<(string, HttpStatusCode statusCode)> GetAccessTokenAsync()
    //    {
    //        var request = new HttpRequestMessage(HttpMethod.Post, _accessTokenUrl)
    //        {
    //            Content = new StringContent("grant_type=client_credentials", Encoding.UTF8, "application/x-www-form-urlencoded")
    //        };

    //        var response = await _httpClient.SendAsync(request);
    //        response.EnsureSuccessStatusCode();

    //        var responseBody = await response.Content.ReadAsStringAsync();
    //        var statusCode = response.StatusCode;
    //        return (responseBody, statusCode);
    //    }

    //    public async Task<(string response, string setCookie, HttpStatusCode statusCode)> PostLoginAsAsync(string url, string requestBody, string accessToken)
    //    {
    //        try
    //        {
    //            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

    //            var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
    //            var response = await _httpClient.PostAsync(url, content);
    //            response.EnsureSuccessStatusCode();

    //            setCookie = response.Headers.GetValues("Set-Cookie").FirstOrDefault();
    //            var responseBody = await response.Content.ReadAsStringAsync();
    //            var statusCode = response.StatusCode;
    //            return (responseBody, setCookie, statusCode);
    //        }
    //        catch (HttpRequestException ex)
    //        {
    //            throw ex;
    //        }
    //        catch (Exception ex)
    //        {
    //            throw ex;
    //        }
    //    }

    //    public async Task<(string response, HttpStatusCode statusCode)> PostGetQueryAsync(string url, string requestBody, string accessToken, string cookieValue)
    //    {
    //        try
    //        {
    //            var uri = new Uri(url);
    //            foreach (var cookie in cookieValue.Split(';'))
    //            {
    //                var cookieParts = cookie.Split('=');
    //                if (cookieParts.Length == 2)
    //                {
    //                    cookieContainer.Add(uri, new Cookie(cookieParts[0].Trim(), cookieParts[1].Trim()));
    //                }
    //            }

    //            //var cookie = new Cookie("");
    //            var reqMsg = new HttpRequestMessage(HttpMethod.Post, url);
    //            reqMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

    //            var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
    //            reqMsg.Content = content;
    //            var response = await _httpClient.SendAsync(reqMsg);
    //            response.EnsureSuccessStatusCode();

    //            var responseBody = await response.Content.ReadAsStringAsync();
    //            var statusCode = response.StatusCode;
    //            return (responseBody, statusCode);
    //        }
    //        catch (HttpRequestException ex)
    //        {
    //            throw ex;
    //        }
    //        catch (Exception ex)
    //        {
    //            throw ex;
    //        }
    //    }

    //    public async Task ExportDocAsync(string url, string requestBody, string accessToken, string cookieValue, string filePath)
    //    {
    //        try
    //        {
    //            var uri = new Uri(url);
    //            foreach (var cookie in cookieValue.Split(';'))
    //            {
    //                var cookieParts = cookie.Split('=');
    //                if (cookieParts.Length == 2)
    //                {
    //                    cookieContainer.Add(uri, new Cookie(cookieParts[0].Trim(), cookieParts[1].Trim()));
    //                }
    //            }

    //            var reqMsg = new HttpRequestMessage(HttpMethod.Get, url);
    //            reqMsg.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

    //            //var content = new StringContent(requestBody, Encoding.UTF8, "application/json");
    //            //reqMsg.Content = content;
    //            var response = await _httpClient.SendAsync(reqMsg);
    //            response.EnsureSuccessStatusCode();

    //            var fileBytes = await response.Content.ReadAsByteArrayAsync();
    //            File.WriteAllBytes(filePath, fileBytes);
    //        }
    //        catch (HttpRequestException ex)
    //        {
    //            throw ex;
    //        }
    //        catch (Exception ex)
    //        {
    //            throw ex;
    //        }
    //    }
    //}
}