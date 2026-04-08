// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Text;

namespace W365ComputerUseSample.ComputerUse;

/// <summary>
/// Sends CUA requests to H Company's Holo3 API (OpenAI Chat Completions compatible).
/// Configure via AIServices:Holo3:ApiKey and AIServices:Holo3:ModelName in appsettings.json.
/// </summary>
public class Holo3ModelProvider : ICuaModelProvider
{
    private readonly HttpClient _httpClient;
    private readonly string _apiKey;
    private readonly ILogger<Holo3ModelProvider> _logger;
    private const string BaseUrl = "https://api.hcompany.ai/v1/chat/completions";

    public string ModelName { get; }

    public Holo3ModelProvider(IHttpClientFactory httpClientFactory, IConfiguration configuration, ILogger<Holo3ModelProvider> logger)
    {
        _httpClient = httpClientFactory.CreateClient("WebClient");
        _logger = logger;
        _apiKey = configuration["AIServices:Holo3:ApiKey"]
            ?? throw new InvalidOperationException("AIServices:Holo3:ApiKey is required.");
        ModelName = configuration["AIServices:Holo3:ModelName"] ?? "holo3-35b-a3b";
    }

    public async Task<string> SendAsync(string requestBody, CancellationToken cancellationToken)
    {
        _logger.LogInformation("Holo3 request to {Url}, model: {Model}", BaseUrl, ModelName);
        using var req = new HttpRequestMessage(HttpMethod.Post, BaseUrl);
        req.Headers.Add("Authorization", $"Bearer {_apiKey}");
        req.Content = new StringContent(requestBody, Encoding.UTF8, "application/json");

        var resp = await _httpClient.SendAsync(req, cancellationToken);
        if (!resp.IsSuccessStatusCode)
        {
            var err = await resp.Content.ReadAsStringAsync(cancellationToken);
            throw new HttpRequestException($"Holo3 returned {resp.StatusCode}: {err}");
        }

        return await resp.Content.ReadAsStringAsync(cancellationToken);
    }
}
