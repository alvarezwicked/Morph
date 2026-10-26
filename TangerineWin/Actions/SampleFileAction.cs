using TangerineWin.Models;

namespace TangerineWin.Actions;

public sealed class SampleFileAction : IFileAction
{
    public string Id => "sample.convert";
    public string DisplayName => "Convert";
    public string Glyph => "⇄";

    public bool CanHandle(IReadOnlyList<string> files) => files.Count > 0;

    public Task<FileActionResult> ExecuteAsync(
        IReadOnlyList<string> files,
        CancellationToken cancellationToken = default)
    {
        // TODO: replace with FFmpeg / ImageMagick / native encoder pipeline.
        return Task.FromResult(new FileActionResult(true));
    }
}
