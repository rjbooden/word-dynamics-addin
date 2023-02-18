using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Web;
using Microsoft.Identity.Web.Resource;
using word_dynamics_api.Helpers;
using word_dynamics_api.Models;

namespace word_dynamics_api.Controllers;

// Simple Snippets settings storing controller
// Based on: https://blog.mastykarz.nl/easiest-store-user-settings-microsoft-365-app/
[Authorize]
[RequiredScope("access_as_user")]
public class SnippetsController : ControllerBase
{
    static readonly string GraphSettingsPath = "me/drive/special/approot:/snippets.json:/content";

    static JsonSerializerOptions jsonOptions = new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private readonly ILogger<SnippetsController> _logger;
    private IHttpClientFactory _httpClient;
    private ITokenAcquisition _tokenAcquisition;
    private IConfiguration _configuration;

    public SnippetsController(ILogger<SnippetsController> logger, IConfiguration configuration, ITokenAcquisition tokenAcquisition, IHttpClientFactory httpClient)
    {
        _logger = logger;
        _httpClient = httpClient;
        _tokenAcquisition = tokenAcquisition;
        _configuration = configuration;
    }

    [HttpGet]
    public async Task<IActionResult> Snippets()
    {
        try {
            var client = _httpClient.CreateClient("msgraph");
            var scope = _configuration.GetValue<string>("GraphScope");
            if (scope == null) {
                throw new Exception("Configuration 'GraphScope' not found");
            }
            var graphEndpoint = _configuration.GetValue<string>("GraphEndpoint");
            if (graphEndpoint == null) {
                throw new Exception("Configuration 'GraphEndpoint' not found");
            } 
            var accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(new [] { scope });
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("text/plain"));
            var query = Request.QueryString.HasValue ? Request.QueryString.Value : "";
            if(!graphEndpoint.EndsWith("/")) {
                graphEndpoint += "/";
            }
            var response = await client.GetAsync($"{graphEndpoint}{GraphSettingsPath}");
            response.Headers.CopyHeaders(this.Response.Headers);
            if (response.StatusCode == HttpStatusCode.OK) {
                var jsonData = await response.Content.ReadAsStringAsync();
                if (string.IsNullOrEmpty(jsonData)) {
                    jsonData = "[]";
                }
                Snippet[]? Snippets = JsonSerializer.Deserialize<Snippet[]>(jsonData, jsonOptions);
                return Ok(Snippets);
            }
            else if (response.StatusCode == HttpStatusCode.NotFound) {
                return Content("[]");
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

    [HttpPost]
    public async Task<IActionResult> Snippets([FromBody] Snippet[] Snippets)
    {
        try {
            var client = _httpClient.CreateClient("msgraph");
            var scope = _configuration.GetValue<string>("GraphScope");
            if (scope == null) {
                throw new Exception("Configuration 'GraphScope' not found");
            }
            var graphEndpoint = _configuration.GetValue<string>("GraphEndpoint");
            if (graphEndpoint == null) {
                throw new Exception("Configuration 'GraphEndpoint' not found");
            } 
            var accessToken = await _tokenAcquisition.GetAccessTokenForUserAsync(new [] { scope });
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            var query = Request.QueryString.HasValue ? Request.QueryString.Value : "";
            if(!graphEndpoint.EndsWith("/")) {
                graphEndpoint += "/";
            }
            var httpContent = new StringContent(JsonSerializer.Serialize<Snippet[]>(Snippets, jsonOptions), Encoding.UTF8, "text/plain");
            var response = await client.PutAsync($"{graphEndpoint}{GraphSettingsPath}", httpContent);
            response.Headers.CopyHeaders(this.Response.Headers);
            if (response.StatusCode == HttpStatusCode.OK) {
                return Ok();
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
