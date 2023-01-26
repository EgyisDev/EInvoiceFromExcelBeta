using EInvoice.API.Services;
using Microsoft.AspNetCore.Mvc;

namespace EInvoice.API.Controllers
{
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly IAuthService _authService;

        public AuthController(IAuthService authService)
        {
            _authService = authService;
        }

        [HttpGet(Name = "GetAccessToken")]
        public async Task<string> GetAccessToken()
        {
            return await _authService.GetAccessToken();
        }
    }
}
