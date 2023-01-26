using System.Text.Json;
using EInvoice.API.Application.EInvoice.DTOs;
using EInvoice.Common.Configurations;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;

namespace EInvoice.API.Services
{
    public class AuthService : IAuthService
    {
        private readonly AuthConfiguration _authConfiguration;
        readonly IMemoryCache _cache;
        public AuthService(IOptions<AuthConfiguration> authConfiguration, IMemoryCache cache)
        {
            _cache = cache;
            _authConfiguration = authConfiguration.Value;
        }
        public async Task<string> GetAccessToken()
        {
            var cacheEntryOptions = new MemoryCacheEntryOptions()
                .SetSlidingExpiration(TimeSpan.FromMinutes(30))
                .SetAbsoluteExpiration(TimeSpan.FromMinutes(50))
                .SetPriority(CacheItemPriority.Normal);

            if (!_cache.TryGetValue("AccessToken", out string accessToken))
            {
                var token = await GetToken();
                accessToken = token;
                _cache.Set("AccessToken", accessToken, cacheEntryOptions);
            }

            return accessToken;
        }

        private async Task<string> GetToken()
        {
            var client = new HttpClient();
            var request = new HttpRequestMessage(HttpMethod.Post, _authConfiguration.TokenUrl);
            var body = new Dictionary<string, string>
            {
                { "client_id", _authConfiguration.ClientId },
                { "client_secret", _authConfiguration.ClientSecret },
                { "scope", _authConfiguration.Scope },
                { "grant_type", _authConfiguration.GrantType }
            };
            request.Content = new FormUrlEncodedContent(body);

            var response = await client.SendAsync(request);

            if (!response.IsSuccessStatusCode)
                throw new Exception("Couldn't be authenticated");

            var responseContent = await response.Content.ReadAsStringAsync();

            var responseJson = JsonSerializer.Deserialize<AccessTokenResult>(responseContent);

            return responseJson.AccessToken;
        }

    }
}
