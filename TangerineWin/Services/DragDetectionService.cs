using System.Windows.Threading;
using TangerineWin.Interop;

namespace TangerineWin.Services;

public sealed class DragDetectionService : IDisposable
{
    private readonly DispatcherTimer _timer;

    private bool _mouseWasDown;
    private bool _isDragging;

    private NativeMethods.POINT _mouseDownPosition;

    private const int DragThreshold = 8;

    public bool IsDragging => _isDragging;

    public event EventHandler? DragStarted;
    public event EventHandler? DragEnded;

    public DragDetectionService()
    {
        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(20)
        };

        _timer.Tick += OnTick;
    }

    public void Start()
    {
        _timer.Start();
    }

    private void OnTick(object? sender, EventArgs e)
    {
        var mouseDown =
            (NativeMethods.GetAsyncKeyState(NativeMethods.VK_LBUTTON) & 0x8000) != 0;

        NativeMethods.GetCursorPos(out var cursor);

        // Только что нажали ЛКМ
        if (mouseDown && !_mouseWasDown)
        {
            _mouseDownPosition = cursor;
            _isDragging = false;
        }

        // ЛКМ удерживается
        if (mouseDown && _mouseWasDown && !_isDragging)
        {
            var dx = Math.Abs(cursor.X - _mouseDownPosition.X);
            var dy = Math.Abs(cursor.Y - _mouseDownPosition.Y);

            if (dx >= DragThreshold || dy >= DragThreshold)
            {
                _isDragging = true;
                DragStarted?.Invoke(this, EventArgs.Empty);
            }
        }

        // ЛКМ отпустили
        if (!mouseDown && _mouseWasDown)
        {
            if (_isDragging)
            {
                _isDragging = false;
                DragEnded?.Invoke(this, EventArgs.Empty);
            }
        }

        _mouseWasDown = mouseDown;
    }

    public void Dispose()
    {
        _timer.Stop();
    }
}