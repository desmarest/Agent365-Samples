// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.AI;

namespace W365ComputerUseSample.ComputerUse;

/// <summary>
/// Lightweight CUA orchestrator for Holo3 models (Chat Completions API).
/// Sends screenshots as base64 images, gets structured action responses,
/// and maps them to W365 MCP tool calls.
/// </summary>
public class Holo3CuaOrchestrator
{
    private readonly Holo3ModelProvider _modelProvider;
    private readonly ILogger<Holo3CuaOrchestrator> _logger;
    private readonly int _maxIterations;

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private const string SystemPrompt = """
        You are a computer-use agent controlling a Windows desktop. You see a screenshot of the current screen.

        Based on the user's task, decide the SINGLE next action to take. Respond with JSON only.

        Available actions:
        - {"action": "click", "x": <int>, "y": <int>, "button": "left|right"}
        - {"action": "double_click", "x": <int>, "y": <int>}
        - {"action": "type", "text": "<string>"}
        - {"action": "keypress", "keys": ["key1", "key2"]}
        - {"action": "scroll", "x": <int>, "y": <int>, "scroll_y": <int>}
        - {"action": "move", "x": <int>, "y": <int>}
        - {"action": "wait", "ms": <int>}
        - {"action": "done", "summary": "<what was accomplished>"}

        Rules:
        - Output ONLY valid JSON, no markdown, no explanation.
        - Examine the screenshot carefully before acting.
        - If you see browser setup or sign-in dialogs, dismiss them.
        - When the task is complete, use the "done" action.
        """;

    public Holo3CuaOrchestrator(
        Holo3ModelProvider modelProvider,
        IConfiguration configuration,
        ILogger<Holo3CuaOrchestrator> logger)
    {
        _modelProvider = modelProvider;
        _logger = logger;
        _maxIterations = configuration.GetValue("ComputerUse:MaxIterations", 30);
    }

    /// <summary>
    /// Run the Holo3 CUA loop. Reuses the same W365 session management as the main orchestrator.
    /// </summary>
    public async Task<string> RunAsync(
        string userMessage,
        string? sessionId,
        IList<AITool> w365Tools,
        Action<string>? onStatusUpdate = null,
        CancellationToken cancellationToken = default)
    {
        _logger.LogInformation("Holo3 CUA loop starting: {Message}", userMessage.Length > 100 ? userMessage[..100] : userMessage);

        // Capture initial screenshot
        var screenshot = await CaptureScreenshotAsync(w365Tools, sessionId, cancellationToken);
        var conversationHistory = new List<object>
        {
            new { role = "system", content = SystemPrompt }
        };

        // Add user task with initial screenshot
        conversationHistory.Add(CreateUserMessage(userMessage, screenshot));

        for (var i = 0; i < _maxIterations; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            // Call Holo3
            var response = await CallModelAsync(conversationHistory, cancellationToken);
            _logger.LogInformation("Holo3 iteration {Iteration}: {Response}", i + 1, response.Length > 200 ? response[..200] : response);

            // Parse action
            Holo3Action? action;
            try
            {
                action = JsonSerializer.Deserialize<Holo3Action>(response, JsonOpts);
            }
            catch (JsonException ex)
            {
                _logger.LogWarning(ex, "Failed to parse Holo3 response as action: {Response}", response);
                // Add response as assistant message and ask to try again
                conversationHistory.Add(new { role = "assistant", content = response });
                conversationHistory.Add(new { role = "user", content = "That was not valid JSON. Please respond with only a JSON action object." });
                continue;
            }

            if (action == null)
                break;

            // Add assistant response to history
            conversationHistory.Add(new { role = "assistant", content = response });

            // Check if done
            if (action.Action == "done")
            {
                _logger.LogInformation("Holo3 task completed: {Summary}", action.Summary);
                return action.Summary ?? "Task completed successfully.";
            }

            // Execute the action
            onStatusUpdate?.Invoke($"Performing: {action.Action}...");
            await ExecuteActionAsync(action, w365Tools, sessionId, cancellationToken);

            // Capture new screenshot
            screenshot = await CaptureScreenshotAsync(w365Tools, sessionId, cancellationToken);

            // Add screenshot as user message for next iteration
            conversationHistory.Add(CreateUserMessage("Here is the screen after the action. What should I do next?", screenshot));
        }

        return "The task could not be completed within the allowed number of steps.";
    }

    private async Task<string> CallModelAsync(List<object> conversation, CancellationToken ct)
    {
        var body = JsonSerializer.Serialize(new
        {
            model = _modelProvider.ModelName,
            messages = conversation,
            temperature = 0.0,
            extra_body = new
            {
                chat_template_kwargs = new { enable_thinking = false }
            }
        }, JsonOpts);

        var responseJson = await _modelProvider.SendAsync(body, ct);

        using var doc = JsonDocument.Parse(responseJson);
        var content = doc.RootElement
            .GetProperty("choices")[0]
            .GetProperty("message")
            .GetProperty("content")
            .GetString() ?? "";

        return content.Trim();
    }

    private static object CreateUserMessage(string text, string base64Screenshot)
    {
        return new
        {
            role = "user",
            content = new object[]
            {
                new { type = "text", text },
                new
                {
                    type = "image_url",
                    image_url = new { url = $"data:image/png;base64,{base64Screenshot}" }
                }
            }
        };
    }

    private async Task ExecuteActionAsync(Holo3Action action, IList<AITool> tools, string? sessionId, CancellationToken ct)
    {
        var (toolName, args) = action.Action switch
        {
            "click" => ("W365_Click2", new Dictionary<string, object?>
            {
                ["x"] = action.X, ["y"] = action.Y, ["button"] = action.Button ?? "left"
            }),
            "double_click" => ("W365_DoubleClick", new Dictionary<string, object?>
            {
                ["x"] = action.X, ["y"] = action.Y
            }),
            "type" => ("W365_WriteText", new Dictionary<string, object?>
            {
                ["text"] = action.Text
            }),
            "keypress" => ("W365_MultiKeyPress", new Dictionary<string, object?>
            {
                ["keys"] = action.Keys ?? Array.Empty<string>()
            }),
            "scroll" => ("W365_Scroll", new Dictionary<string, object?>
            {
                ["atX"] = action.X, ["atY"] = action.Y, ["deltaX"] = 0, ["deltaY"] = action.ScrollY ?? 0
            }),
            "move" => ("W365_MoveMouse", new Dictionary<string, object?>
            {
                ["toX"] = action.X, ["toY"] = action.Y
            }),
            "wait" => ("W365_Wait", new Dictionary<string, object?>
            {
                ["milliseconds"] = action.Ms ?? 500
            }),
            _ => throw new NotSupportedException($"Unsupported Holo3 action: {action.Action}")
        };

        if (!string.IsNullOrEmpty(sessionId))
            args["sessionId"] = sessionId;

        await ComputerUseOrchestrator.InvokeToolAsync(tools, toolName, args, ct);
    }

    private static async Task<string> CaptureScreenshotAsync(IList<AITool> tools, string? sessionId, CancellationToken ct)
    {
        var args = new Dictionary<string, object?>();
        if (!string.IsNullOrEmpty(sessionId))
            args["sessionId"] = sessionId;

        var result = await ComputerUseOrchestrator.InvokeToolAsync(tools, "W365_CaptureScreenshot", args, ct);
        var str = result?.ToString() ?? "";

        // Try to extract base64 image data from the response
        try
        {
            using var doc = JsonDocument.Parse(str);
            var root = doc.RootElement;
            if (root.TryGetProperty("screenshotData", out var sd)) return sd.GetString() ?? "";
            if (root.TryGetProperty("image", out var img)) return img.GetString() ?? "";
            if (root.TryGetProperty("data", out var d)) return d.GetString() ?? "";
            if (root.TryGetProperty("content", out var content) && content.ValueKind == JsonValueKind.Array)
            {
                foreach (var block in content.EnumerateArray())
                {
                    if (block.TryGetProperty("data", out var blockData) && !string.IsNullOrEmpty(blockData.GetString()))
                        return blockData.GetString()!;
                }
            }
        }
        catch (JsonException) { }

        if (str.Length > 1000 && !str.StartsWith("{") && !str.StartsWith("["))
            return str;

        throw new InvalidOperationException($"Holo3: Failed to extract screenshot. Response length: {str.Length}");
    }

    /// <summary>
    /// Structured action response from Holo3.
    /// </summary>
    private class Holo3Action
    {
        [JsonPropertyName("action")]
        public string Action { get; set; } = "";

        [JsonPropertyName("x")]
        public int? X { get; set; }

        [JsonPropertyName("y")]
        public int? Y { get; set; }

        [JsonPropertyName("button")]
        public string? Button { get; set; }

        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("keys")]
        public string[]? Keys { get; set; }

        [JsonPropertyName("scroll_y")]
        public int? ScrollY { get; set; }

        [JsonPropertyName("ms")]
        public int? Ms { get; set; }

        [JsonPropertyName("summary")]
        public string? Summary { get; set; }
    }
}
