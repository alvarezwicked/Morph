using System.Windows;
using TangerineWin.Services;
using TangerineWin.Views;

namespace TangerineWin;

public partial class App : System.Windows.Application
{
    private TrayService? _trayService;
    private HotkeyService? _hotkeyService;
    private OverlayWindow? _overlay;
    private DragDetectionService? _dragDetectionService;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _dragDetectionService = new DragDetectionService();
        _overlay = new OverlayWindow();
        _overlay.DragFilesDetected += files =>
        {
            System.Diagnostics.Debug.WriteLine(
                "Dragged files: " + string.Join(", ", files));
        };
        _trayService = new TrayService(
            onShow: () => _overlay.ShowNearCursor(),
            onExit: Shutdown);

        _hotkeyService = new HotkeyService();
        _hotkeyService.ShiftPressed += (_, _) =>
        {
            if (!_dragDetectionService.IsDragging)
                return;

            _overlay.ShowNearCursor();
        };

        _hotkeyService.ShiftReleased += (_, _) =>
        {
            _overlay.Hide();
        };
        _dragDetectionService.DragEnded += (_, _) =>
        {
            _overlay.Hide();
        };
        _hotkeyService.Start();
        _dragDetectionService.Start();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _hotkeyService?.Dispose();
        _dragDetectionService?.Dispose();
        _trayService?.Dispose();

        base.OnExit(e);
    }
}
