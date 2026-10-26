using System.Windows.Threading;
using TangerineWin.Interop;

namespace TangerineWin.Services;

public sealed class HotkeyService : IDisposable
{
    private readonly DispatcherTimer _timer;
    private bool _wasDown;

    public event EventHandler? ShiftPressed;
    public event EventHandler? ShiftReleased;

    public HotkeyService()
    {
        _timer = new DispatcherTimer
        {
            Interval = TimeSpan.FromMilliseconds(40)
        };
        _timer.Tick += OnTick;
    }

    public void Start() => _timer.Start();

    private void OnTick(object? sender, EventArgs e)
    {
        var isDown =
            (NativeMethods.GetAsyncKeyState(NativeMethods.VK_SHIFT) & 0x8000) != 0;

        if (isDown && !_wasDown)
        {
            ShiftPressed?.Invoke(this, EventArgs.Empty);
        }
        else if (!isDown && _wasDown)
        {
            ShiftReleased?.Invoke(this, EventArgs.Empty);
        }

        _wasDown = isDown;
    }

    public void Dispose() => _timer.Stop();
}
