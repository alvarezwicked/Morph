using TangerineWin.Models;

namespace TangerineWin.Actions;

public interface IFileAction
{
    string Id { get; }
    string DisplayName { get; }
    string Glyph { get; }

    bool CanHandle(IReadOnlyList<string> files);

    Task<FileActionResult> ExecuteAsync(
        IReadOnlyList<string> files,
        CancellationToken cancellationToken = default);
}
