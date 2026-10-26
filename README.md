# TangerineWin Starter

Минимальная заготовка Windows-приложения в духе Tangerine для macOS.

## Что уже есть

- .NET 8 + WPF
- tray-приложение
- прозрачное topmost overlay-окно
- radial-menu заглушка
- глобальная реакция на нажатие `Shift`
- позиционирование меню около текущего курсора
- интерфейс `IFileAction`
- `ActionRegistry` для будущих конвертеров

## Что пока НЕ реализовано

Самая важная часть оставлена следующим этапом:

1. определение активного OLE drag-and-drop из Windows Explorer;
2. получение путей файлов из текущего drag session;
3. показ меню только во время реального drag;
4. корректный click-through / interaction mode overlay-окна;
5. выбор сектора radial menu движением курсора;
6. FFmpeg / ImageMagick / PDF processing;
7. progress UI, settings, autostart, updater, installer.

## Запуск

Нужны Windows 10/11 и .NET 8 SDK.

```powershell
dotnet restore
dotnet run --project .\TangerineWin\TangerineWin.csproj
```

После запуска приложение появится в system tray.

- Нажмите `Shift` — radial overlay появится около курсора.
- Двойной клик по tray icon делает то же самое.

Это пока proof-of-concept UI/architecture, а не полноценный перехват drag из Explorer.

## Следующий технический шаг

Рекомендуемый следующий milestone — `ExplorerDragMonitor`:

```text
Explorer/OLE drag
       ↓
ExplorerDragMonitor
       ↓
DraggedFilesChanged(files)
       ↓
ActionRegistry
       ↓
OverlayWindow.ShowNearCursor(actions)
```

Для production-версии лучше уйти от polling клавиши к более низкоуровневому Win32/OLE механизму и аккуратно работать с COM drag/drop.
