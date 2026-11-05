namespace Morph.Models;

public sealed record FileActionResult(bool Success, string? OutputPath = null, string? Error = null);
