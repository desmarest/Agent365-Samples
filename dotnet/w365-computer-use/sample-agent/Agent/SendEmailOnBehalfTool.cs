// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.ComponentModel;
using System.IdentityModel.Tokens.Jwt;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Identity.Abstractions;
using Microsoft.Identity.Web;

namespace W365ComputerUseSample.Agent;

/// <summary>
/// Builds the <c>send_email_on_behalf</c> AITool: reads a pre-minted human Tc from
/// configuration (<c>Obo:UserAccessToken</c>), exchanges it for a Graph TR via agent-OBO
/// through the agent identity using <c>Microsoft.Identity.Web.AgentIdentities</c>, and
/// calls <c>/me/sendMail</c> as the human user. The pasted-Tc shortcut is used because
/// the Agents SDK's <c>AzureBotUserAuthorization</c> handler is incompatible with an
/// agentic-marked bot identity (blueprint app cannot do client_credentials to BF).
/// </summary>
public static class SendEmailOnBehalfTool
{
    public static AIFunction Create(
        string userAccessToken,
        IAuthorizationHeaderProvider headerProvider,
        IHttpClientFactory httpClientFactory,
        string agentIdentityAppId,
        ILogger logger)
    {
        return AIFunctionFactory.Create(
            async (
                [Description("Recipient email address.")] string to,
                [Description("Subject line.")] string subject,
                [Description("Plain-text body.")] string body,
                CancellationToken cancellationToken) =>
            {
                logger.LogWarning("send_email_on_behalf: INVOKED — to={To}, subject={Subject}", to, subject);

                if (string.IsNullOrEmpty(userAccessToken))
                {
                    logger.LogWarning("OBO POC: Obo:UserAccessToken is not set.");
                    return "No user access token (Tc) is configured. Set Obo:UserAccessToken in app settings.";
                }

                logger.LogWarning("OBO POC: Tc acquired ({Len} chars). Exchanging for TR.", userAccessToken.Length);

                ClaimsPrincipal userPrincipal = BuildClaimsPrincipalFromJwt(userAccessToken);

                var options = new AuthorizationHeaderProviderOptions()
                    .WithAgentIdentity(agentIdentityAppId);

                string authHeader;
                try
                {
                    authHeader = await headerProvider.CreateAuthorizationHeaderForUserAsync(
                        scopes: new[] { "https://graph.microsoft.com/Mail.Send" },
                        authorizationHeaderProviderOptions: options,
                        claimsPrincipal: userPrincipal,
                        cancellationToken: cancellationToken).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    logger.LogError(ex, "OBO POC: agent-OBO exchange failed.");
                    return $"Failed to acquire delegated Graph token: {ex.Message}";
                }

                logger.LogWarning("OBO POC: TR acquired ({Len} chars). Posting to Graph.", authHeader.Length);

                var payload = new
                {
                    message = new
                    {
                        subject,
                        body = new { contentType = "Text", content = body },
                        toRecipients = new[]
                        {
                            new { emailAddress = new { address = to } },
                        },
                    },
                    saveToSentItems = true,
                };

                using HttpClient http = httpClientFactory.CreateClient("WebClient");
                using var request = new HttpRequestMessage(HttpMethod.Post, "https://graph.microsoft.com/v1.0/me/sendMail")
                {
                    Content = new StringContent(
                        JsonSerializer.Serialize(payload),
                        Encoding.UTF8,
                        "application/json"),
                };
                request.Headers.Authorization = AuthenticationHeaderValue.Parse(authHeader);

                HttpResponseMessage response = await http.SendAsync(request, cancellationToken).ConfigureAwait(false);
                if (response.IsSuccessStatusCode)
                {
                    logger.LogWarning("OBO POC: sent mail to {To} on behalf of human user.", to);
                    return $"Sent email to {to} on your behalf.";
                }

                string errorBody = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
                logger.LogWarning("OBO POC: Graph returned {Status}: {Body}", (int)response.StatusCode, errorBody);
                return $"Graph sendMail failed ({(int)response.StatusCode}): {errorBody}";
            },
            name: "send_email_on_behalf",
            description: "Sends an email from the signed-in human user's mailbox (not the agent user's). Use this when the user explicitly says 'on my behalf' or asks to send mail as themselves.");
    }

    private static ClaimsPrincipal BuildClaimsPrincipalFromJwt(string rawJwt)
    {
        var handler = new JwtSecurityTokenHandler();
        JwtSecurityToken parsed = handler.ReadJwtToken(rawJwt);

        var identity = new ClaimsIdentity(parsed.Claims, authenticationType: "Bearer")
        {
            BootstrapContext = rawJwt,
        };

        return new ClaimsPrincipal(identity);
    }
}
