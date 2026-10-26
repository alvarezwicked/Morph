using System.Runtime.InteropServices;
using TangerineWin.Interop;
using System.IO;
namespace TangerineWin.Services;

public sealed class ExplorerSelectionService
{
    public IReadOnlyList<string> GetSelectedFiles()
    {
        var result = new List<string>();

        try
        {
            var foregroundWindow = NativeMethods.GetForegroundWindow();

            var shellType = Type.GetTypeFromProgID("Shell.Application");
            if (shellType is null)
                return result;

            dynamic? shell = Activator.CreateInstance(shellType);

            if (shell is null)
                return result;

            dynamic windows = shell.Windows();

            foreach (var window in windows)
            {
                try
                {
                    var hwnd = new IntPtr((long)window.HWND);

                    // Нас интересует только активное окно Explorer
                    if (hwnd != foregroundWindow)
                        continue;

                    dynamic document = window.Document;
                    dynamic selectedItems = document.SelectedItems();

                    foreach (var item in selectedItems)
                    {
                        string? path = item.Path as string;

                        if (string.IsNullOrWhiteSpace(path))
                            continue;

                        if (File.Exists(path))
                        {
                            result.Add(path);
                        }
                    }

                    break;
                }
                catch
                {
                    // Это может быть не обычное окно Explorer
                }
            }
        }
        catch
        {
            // Explorer может быть недоступен — просто считаем,
            // что ничего не выбрано.
        }

        return result;
    }
}