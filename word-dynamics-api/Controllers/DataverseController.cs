using System.Net.Http.Headers;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;
using word_dynamics_api.Helpers;

namespace word_dynamics_api.Controllers;

// Dataverse simple get proxy
[Authorize]
[RequiredScope("access_as_user")]
public class DataverseController : ControllerBase
{

    private readonly ILogger<DataverseController> _logger;
    private IHttpClientFactory _httpClient;
    private ITokenAcquisition _tokenAcquisition;
    private IConfiguration _configuration;

    public DataverseController(ILogger<DataverseController> logger, IConfiguration configuration, ITokenAcquisition tokenAcquisition, IHttpClientFactory httpClient)
    {
        _logger = logger;
        _httpClient = httpClient;
        _tokenAcquisition = tokenAcquisition;
        _configuration = configuration;
    }

    [HttpGet]
    public async Task<IActionResult> Api(string id)
    {
        try {
            var client = _httpClient.CreateClient("dataverse");
            var scope = _configuration.GetValue<string>("DataverseScope");
            if (scope == null) {
                throw new Exception("Configuration 'DataverseScope' not found");
            }
            var dataverseEndpoint = _configuration.GetValue<string>("DataverseEndpoint");
            if (dataverseEndpoint == null) {
                throw new Exception("Configuration 'DataverseEndpoint' not found");
            } 
            var accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(new [] { scope });
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            var query = Request.QueryString.HasValue ? Request.QueryString.Value : "";
            if(!dataverseEndpoint.EndsWith("/")) {
                dataverseEndpoint += "/";
            }
            var response = await client.GetAsync($"{dataverseEndpoint}{id}{query}");
            response.Headers.CopyHeaders(this.Response.Headers);
            if (response.StatusCode == System.Net.HttpStatusCode.OK) {
                var jsonData = await response.Content.ReadAsStringAsync();
                return Content(jsonData);
            }
            return StatusCode((int)response.StatusCode, response.ReasonPhrase);
        }
        catch(MicrosoftIdentityWebChallengeUserException miex) {
            if (null != miex.MsalUiRequiredException) {
                return StatusCode(401, miex.MsalUiRequiredException.Message);
            }
            return StatusCode(401, miex.Message);
        }
        catch(Exception ex) {
            _logger.LogError(500, ex, "Error in Dynamics get proxy");
            return StatusCode(500);
        }
    }

}
