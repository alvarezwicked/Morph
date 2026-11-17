
using System;
using System.Threading;
using System.Windows;
using Morph.Services;
using Morph.Views;

namespace Morph;

public partial class App : System.Windows.Application
{
    private const string MutexName =
        @"Local\vincente.soft.Morph.SingleInstance";

    private Mutex? _instanceMutex;
    private bool _ownsMutex;

    private TrayService? _trayService;
    private HotkeyService? _hotkeyService;
    private OverlayWindow? _overlay;
    private DragDetectionService? _dragDetectionService;

    protected override void OnStartup(StartupEventArgs e)
    {
        try
        {
            _instanceMutex = new Mutex(
                initiallyOwned: true,
                name: MutexName,
                createdNew: out bool createdNew);

            _ownsMutex = createdNew;

            if (!createdNew)
            {
                _instanceMutex.Dispose();
                _instanceMutex = null;

                Shutdown();
                return;
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                $"Failed to initialize Morph:\n{ex.Message}",
                "Morph",
                MessageBoxButton.OK,
                MessageBoxImage.Error);

            Shutdown();
            return;
        }

        base.OnStartup(e);

        _dragDetectionService =
            new DragDetectionService();

        _overlay =
            new OverlayWindow();

        _overlay.DragFilesDetected += files =>
        {
            System.Diagnostics.Debug.WriteLine(
                "Dragged files: " + string.Join(", ", files));
        };

        _trayService = new TrayService(
            onShow: () => _overlay.ShowNearCursor(),
            onExit: Shutdown);

        // Получаем уведомления из OverlayWindow
        // и показываем их через системный трей.
        _overlay.NotificationRequested +=
            (title, message, isError) =>
            {
                _trayService?.ShowNotification(
                    title,
                    message,
                    isError);
            };

        _hotkeyService =
            new HotkeyService();

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

        if (_ownsMutex && _instanceMutex is not null)
        {
            try
            {
                _instanceMutex.ReleaseMutex();
            }
            catch (ApplicationException)
            {
            }
        }

        _instanceMutex?.Dispose();
        _instanceMutex = null;
        _ownsMutex = false;

        base.OnExit(e);
    }
}
