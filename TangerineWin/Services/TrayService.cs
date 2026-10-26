using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace TangerineWin.Services;

public sealed class TrayService : IDisposable
{
    private readonly NotifyIcon _notifyIcon;
    private readonly Icon? _applicationIcon;

    public TrayService(
        Action onShow,
        Action onExit)
    {
        var menu =
            new ContextMenuStrip();

        menu.Items.Add(
            "Show",
            null,
            (_, _) => onShow());

        menu.Items.Add(
            "Exit",
            null,
            (_, _) => onExit());

        var executablePath =
            Environment.ProcessPath;

        if (!string.IsNullOrWhiteSpace(
                executablePath) &&
            File.Exists(
                executablePath))
        {
            _applicationIcon =
                Icon.ExtractAssociatedIcon(
                    executablePath!);
        }

        _notifyIcon =
            new NotifyIcon
            {
                Text =
                    "Morph",

                Icon =
                    _applicationIcon ??
                    SystemIcons.Application,

                Visible =
                    true,

                ContextMenuStrip =
                    menu
            };

        _notifyIcon.DoubleClick +=
            (_, _) => onShow();
    }

    public void Dispose()
    {
        _notifyIcon.Visible =
            false;

        _notifyIcon.Dispose();

        _applicationIcon?.Dispose();
    }
}