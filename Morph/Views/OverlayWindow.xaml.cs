using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;

using Morph.Interop;
using Morph.Models;
using Morph.Services;

using WpfBorder = System.Windows.Controls.Border;
using WpfButton = System.Windows.Controls.Button;
using WpfCanvas = System.Windows.Controls.Canvas;
using WpfGrid = System.Windows.Controls.Grid;
using WpfOrientation = System.Windows.Controls.Orientation;
using WpfStackPanel = System.Windows.Controls.StackPanel;
using WpfTextBlock = System.Windows.Controls.TextBlock;

using WpfBrush = System.Windows.Media.Brush;
using WpfBrushes = System.Windows.Media.Brushes;
using WpfColor = System.Windows.Media.Color;
using WpfFontFamily = System.Windows.Media.FontFamily;
using WpfGeometry = System.Windows.Media.Geometry;
using WpfGradientStop = System.Windows.Media.GradientStop;
using WpfGradientStopCollection = System.Windows.Media.GradientStopCollection;
using WpfLinearGradientBrush = System.Windows.Media.LinearGradientBrush;
using WpfScaleTransform = System.Windows.Media.ScaleTransform;
using WpfVisualTreeHelper = System.Windows.Media.VisualTreeHelper;

using WpfLine = System.Windows.Shapes.Line;
using WpfPath = System.Windows.Shapes.Path;

using WpfDragDropEffects = System.Windows.DragDropEffects;
using WpfHorizontalAlignment = System.Windows.HorizontalAlignment;
using WpfMessageBox = System.Windows.MessageBox;
using WpfPoint = System.Windows.Point;
using WpfVerticalAlignment = System.Windows.VerticalAlignment;

using WinFormsScreen = System.Windows.Forms.Screen;
using DrawingPoint = System.Drawing.Point;

namespace Morph.Views;

public partial class OverlayWindow : Window
{
    private readonly ActionRegistry _actionRegistry;

    private readonly ImageConversionService
        _imageConversionService;

    private readonly ZipService
        _zipService;

    private readonly PdfConversionService
        _pdfConversionService;

    private readonly OfficeConversionService
        _officeConversionService;

    private readonly PdfExtractionService
        _pdfExtractionService;

    private readonly PdfRenderService
        _pdfRenderService;

    private readonly PdfMergeService
        _pdfMergeService;

    private readonly PdfSplitService
        _pdfSplitService;

    private readonly PdfImageExtractionService
        _pdfImageExtractionService;

    private readonly PdfPageExtractionService
        _pdfPageExtractionService;

    private readonly PdfPageDeletionService
        _pdfPageDeletionService;

    private readonly PdfPageRotationService
        _pdfPageRotationService;

    private readonly PdfPageReorderService
        _pdfPageReorderService;

    private readonly PdfOcrDocxService
        _pdfOcrDocxService;

    private readonly PdfExcelExtractionService
        _pdfExcelExtractionService;

    private readonly PdfPowerPointService
        _pdfPowerPointService;

    private readonly PdfEditablePowerPointService
        _pdfEditablePowerPointService;

    private WpfButton? _activeActionButton;

    private bool _isShowingOverlay;

    public IReadOnlyList<string> DraggedFiles
    {
        get;
        private set;
    } = Array.Empty<string>();


    // Событие обнаружения перетаскиваемых файлов.
    public event Action<IReadOnlyList<string>>?
        DragFilesDetected;


    // Событие для системных уведомлений Windows.
    public event Action<string, string, bool>?
        NotificationRequested;


    // ====================================================
    // CONSTRUCTOR
    // ====================================================

    public OverlayWindow()
    {
        InitializeComponent();

        _actionRegistry =
            new ActionRegistry();

        _imageConversionService =
            new ImageConversionService();

        _zipService =
            new ZipService();

        _pdfConversionService =
            new PdfConversionService();

        _officeConversionService =
            new OfficeConversionService();

        _pdfExtractionService =
            new PdfExtractionService();

        _pdfRenderService =
            new PdfRenderService();

        _pdfMergeService =
            new PdfMergeService();

        _pdfSplitService =
            new PdfSplitService();

        _pdfImageExtractionService =
            new PdfImageExtractionService();

        _pdfPageExtractionService =
            new PdfPageExtractionService();

        _pdfPageDeletionService =
            new PdfPageDeletionService();

        _pdfPageRotationService =
            new PdfPageRotationService();

        _pdfPageReorderService =
            new PdfPageReorderService();

        _pdfOcrDocxService =
            new PdfOcrDocxService();

        _pdfExcelExtractionService =
            new PdfExcelExtractionService();

        _pdfPowerPointService =
            new PdfPowerPointService();

        _pdfEditablePowerPointService =
            new PdfEditablePowerPointService();

        IsVisibleChanged +=
            OnOverlayVisibilityChanged;

        ResetUi();
    }


    // ====================================================
    // SHOW
    // ====================================================

    public void ShowNearCursor()
    {
        //
        // Не запускаем повторный показ,
        // если окно уже видно или сейчас открывается.
        //
        if (_isShowingOverlay ||
            IsVisible)
        {
            return;
        }

        _isShowingOverlay =
            true;

        try
        {
            DraggedFiles =
                Array.Empty<string>();

            ResetUi();

            if (!NativeMethods.GetCursorPos(
                    out var cursorPoint))
            {
                return;
            }

            //
            // Создаём HWND, но НЕ показываем окно.
            //
            // Это ключевой момент для Windows 11:
            // окно сначала будет правильно позиционировано,
            // и только потом станет видимым.
            //
            var helper =
                new WindowInteropHelper(
                    this);

            var hwnd =
                helper.EnsureHandle();

            if (hwnd ==
                IntPtr.Zero)
            {
                return;
            }

            UpdateLayout();

            //
            // Получаем реальный размер окна
            // в физических пикселях.
            //
            if (!GetWindowRect(
                    hwnd,
                    out var windowRect))
            {
                return;
            }

            var windowWidth =
                windowRect.Right -
                windowRect.Left;

            var windowHeight =
                windowRect.Bottom -
                windowRect.Top;

            //
            // Определяем монитор, на котором
            // сейчас находится курсор.
            //
            var screen =
                WinFormsScreen.FromPoint(
                    new DrawingPoint(
                        cursorPoint.X,
                        cursorPoint.Y));

            var workArea =
                screen.WorkingArea;

            const int margin =
                10;

            //
            // Центрируем интерфейс относительно курсора.
            //
            var targetLeft =
                cursorPoint.X -
                windowWidth / 2;

            var targetTop =
                cursorPoint.Y -
                windowHeight / 2;

            //
            // Ограничиваем координаты рабочей областью.
            //
            var minimumLeft =
                workArea.Left +
                margin;

            var minimumTop =
                workArea.Top +
                margin;

            var maximumLeft =
                workArea.Right -
                windowWidth -
                margin;

            var maximumTop =
                workArea.Bottom -
                windowHeight -
                margin;

            if (maximumLeft >=
                minimumLeft)
            {
                targetLeft =
                    Math.Clamp(
                        targetLeft,
                        minimumLeft,
                        maximumLeft);
            }
            else
            {
                targetLeft =
                    workArea.Left;
            }

            if (maximumTop >=
                minimumTop)
            {
                targetTop =
                    Math.Clamp(
                        targetTop,
                        minimumTop,
                        maximumTop);
            }
            else
            {
                targetTop =
                    workArea.Top;
            }

            //
            // Сначала физически ставим скрытое окно
            // на правильное место.
            //
            SetWindowPos(
                hwnd,
                IntPtr.Zero,
                targetLeft,
                targetTop,
                0,
                0,
                SWP_NOSIZE |
                SWP_NOZORDER |
                SWP_NOACTIVATE);

            //
            // Подготавливаем первый кадр анимации
            // ДО Show().
            //
            PrepareOpenAnimationState();

            //
            // Только теперь показываем уже
            // правильно расположенное окно.
            //
            Show();

            //
            // И запускаем анимацию явно.
            //
            PlayOpenAnimation();
        }
        finally
        {
            _isShowingOverlay =
                false;
        }
    }
    // ====================================================
    // BUTTON COLLECTION
    // ====================================================

    private WpfButton[] GetActionButtons()
    {
        return new[]
        {
            Action1,
            Action2,
            Action3,
            Action4,
            Action5,
            Action6,
            Action7,
            Action8,
            Action9,
            Action10,
            Action11,
            Action12,
            Action13,
            Action14
        };
    }


    // ====================================================
    // RADIAL LAYOUT
    // ====================================================

    private void ArrangeActionButtons()
    {
        var buttons =
            GetActionButtons()
                .Where(button =>
                    button.Visibility ==
                    Visibility.Visible &&
                    button.Tag is FileAction)
                .ToArray();

        ConnectorCanvas.Children.Clear();

        if (buttons.Length == 0)
        {
            return;
        }

        const double centerX =
            310;

        const double centerY =
            310;

        const double buttonSize =
            92;

        //
        // Для небольшого количества действий
        // используем отдельные компактные композиции.
        //
        var layout =
            GetAdaptiveActionLayout(
                buttons.Length);

        for (var i = 0;
             i < buttons.Length;
             i++)
        {
            var angleDegrees =
                layout[i].Angle;

            var radius =
                layout[i].Radius;

            var angleRadians =
                angleDegrees *
                Math.PI /
                180.0;

            var directionX =
                Math.Cos(
                    angleRadians);

            var directionY =
                Math.Sin(
                    angleRadians);

            var buttonCenterX =
                centerX +
                directionX *
                radius;

            var buttonCenterY =
                centerY +
                directionY *
                radius;

            WpfCanvas.SetLeft(
                buttons[i],
                buttonCenterX -
                buttonSize / 2.0);

            WpfCanvas.SetTop(
                buttons[i],
                buttonCenterY -
                buttonSize / 2.0);

            AddOrbitConnector(
                centerX,
                centerY,
                directionX,
                directionY,
                radius,
                buttonSize);
        }
    }

    private static (double Angle, double Radius)[]
    GetAdaptiveActionLayout(
        int count)
    {
        switch (count)
        {
            case 1:
                return new[]
                {
                (
                    Angle: -90.0,
                    Radius: 158.0
                )
            };

            case 2:
                return new[]
                {
                (
                    Angle: -130.0,
                    Radius: 168.0
                ),
                (
                    Angle: -50.0,
                    Radius: 168.0
                )
            };

            case 3:
                return new[]
                {
                (
                    Angle: -90.0,
                    Radius: 172.0
                ),
                (
                    Angle: 30.0,
                    Radius: 172.0
                ),
                (
                    Angle: 150.0,
                    Radius: 172.0
                )
            };

            case 4:
                return new[]
                {
                (
                    Angle: -135.0,
                    Radius: 176.0
                ),
                (
                    Angle: -45.0,
                    Radius: 176.0
                ),
                (
                    Angle: 45.0,
                    Radius: 176.0
                ),
                (
                    Angle: 135.0,
                    Radius: 176.0
                )
            };

            case 5:
                return CreateCircularLayout(
                    count,
                    180);

            case 6:
                return CreateCircularLayout(
                    count,
                    184);

            case 7:
                return CreateCircularLayout(
                    count,
                    188);

            case 8:
                return CreateCircularLayout(
                    count,
                    191);

            case 9:
                return CreateCircularLayout(
                    count,
                    194);

            case 10:
                return CreateCircularLayout(
                    count,
                    197);

            case 11:
                return CreateCircularLayout(
                    count,
                    200);

            case 12:
                return CreateCircularLayout(
                    count,
                    202);

            case 13:
                return CreateCircularLayout(
                    count,
                    205);

            default:
                return CreateCircularLayout(
                    count,
                    208);
        }
    }

    private static (double Angle, double Radius)[]
    CreateCircularLayout(
        int count,
        double radius)
    {
        var result =
            new (double Angle, double Radius)[count];

        for (var i = 0;
             i < count;
             i++)
        {
            result[i] =
                (
                    Angle:
                        -90.0 +
                        360.0 *
                        i /
                        count,

                    Radius:
                        radius
                );
        }

        return result;
    }


    private void AddOrbitConnector(
        double centerX,
        double centerY,
        double directionX,
        double directionY,
        double radius,
        double buttonSize)
    {
        const double centerDiskRadius =
            106;

        var buttonRadius =
            buttonSize /
            2.0;

        var startRadius =
            centerDiskRadius +
            1;

        var endRadius =
            radius -
            buttonRadius -
            3;

        if (endRadius <=
            startRadius)
        {
            return;
        }

        var line =
            new WpfLine
            {
                X1 =
                    centerX +
                    directionX *
                    startRadius,

                Y1 =
                    centerY +
                    directionY *
                    startRadius,

                X2 =
                    centerX +
                    directionX *
                    endRadius,

                Y2 =
                    centerY +
                    directionY *
                    endRadius,

                Stroke =
                    new SolidColorBrush(
                        WpfColor.FromArgb(
                            55,
                            196,
                            205,
                            216)),

                StrokeThickness =
                    1,

                StrokeStartLineCap =
                    PenLineCap.Round,

                StrokeEndLineCap =
                    PenLineCap.Round,

                IsHitTestVisible =
                    false
            };

        ConnectorCanvas.Children.Add(
            line);
    }


    // ====================================================
    // ACTION BUTTON CONTENT
    // ====================================================

    private void RefreshActionButtonContents()
    {
        foreach (var button in
                 GetActionButtons())
        {
            if (button.Tag
                is not FileAction action)
            {
                continue;
            }

            button.Content =
                BuildActionButtonContent(
                    action);

            button.Background =
                CreateNormalButtonBrush();

            button.BorderBrush =
                CreateNormalBorderBrush();

            button.BorderThickness =
                new Thickness(
                    1);

            button.Opacity =
                1.0;
        }
    }


    private FrameworkElement
        BuildActionButtonContent(
            FileAction action)
    {
        var icon =
            BuildActionIcon(
                action);

        var title =
            new WpfTextBlock
            {
                Text =
                    GetActionTitle(
                        action),

                FontFamily =
                    new WpfFontFamily(
                        "Segoe UI"),

                FontSize =
                    8.9,

                FontWeight =
                    FontWeights.SemiBold,

                Foreground =
                new SolidColorBrush(
                    WpfColor.FromRgb(
                        33,
                        52,
                        76)),

                TextAlignment =
                    TextAlignment.Center,

                TextWrapping =
                    TextWrapping.Wrap,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                MaxWidth =
                    78,

                Margin =
                    new Thickness(
                        0,
                        5,
                        0,
                        0)
            };

        var subtitle =
            new WpfTextBlock
            {
                Text =
                    GetActionSubtitle(
                        action),

                FontFamily =
                    new WpfFontFamily(
                        "Segoe UI"),

                FontSize =
                    6.9,

                FontWeight =
                    FontWeights.Normal,

                Foreground =
                new SolidColorBrush(
                    WpfColor.FromRgb(
                        86,
                        105,
                        128)),

                TextAlignment =
                    TextAlignment.Center,

                TextWrapping =
                    TextWrapping.Wrap,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                MaxWidth =
                    80,

                Margin =
                    new Thickness(
                        0,
                        1,
                        0,
                        0)
            };

        var panel =
            new WpfStackPanel
            {
                Orientation =
                    WpfOrientation.Vertical,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                VerticalAlignment =
                    WpfVerticalAlignment.Center
            };

        panel.Children.Add(
            icon);

        panel.Children.Add(
            title);

        panel.Children.Add(
            subtitle);

        return panel;
    }


    // ====================================================
    // ICON BUILDER
    // ====================================================

    private FrameworkElement BuildActionIcon(
        FileAction action)
    {
        return action.Id switch
        {
            "pdf_to_docx" =>
                BuildOfficeIcon(
                    "W",
                    WpfColor.FromRgb(
                        36,
                        103,
                        190)),

            "pdf_to_xlsx" =>
                BuildOfficeIcon(
                    "X",
                    WpfColor.FromRgb(
                        33,
                        129,
                        92)),

            "pdf_to_pptx" =>
                BuildOfficeIcon(
                    "P",
                    WpfColor.FromRgb(
                        210,
                        72,
                        54)),

            "pdf_to_ocr_docx" =>
                BuildVectorIcon(
                    GetScanGeometry(),
                    WpfColor.FromRgb(
                        45,
                        143,
                        141)),

            "pdf_to_png" =>
                BuildVectorIcon(
                    GetImageGeometry(),
                    WpfColor.FromRgb(
                        104,
                        93,
                        215)),

            "pdf_to_jpg" =>
                BuildVectorIcon(
                    GetImageGeometry(),
                    WpfColor.FromRgb(
                        184,
                        91,
                        138)),

            "extract_pdf_images" =>
                BuildVectorIcon(
                    GetImageGeometry(),
                    WpfColor.FromRgb(
                        214,
                        157,
                        59)),

            "split_pdf" =>
                BuildVectorIcon(
                    GetScissorsGeometry(),
                    WpfColor.FromRgb(
                        116,
                        82,
                        206)),

            "rotate_pdf_pages" =>
                BuildVectorIcon(
                    GetRotateGeometry(),
                    WpfColor.FromRgb(
                        56,
                        125,
                        212)),

            "delete_pdf_pages" =>
                BuildVectorIcon(
                    GetTrashGeometry(),
                    WpfColor.FromRgb(
                        215,
                        67,
                        82)),

            "reorder_pdf_pages" =>
                BuildVectorIcon(
                    GetReorderGeometry(),
                    WpfColor.FromRgb(
                        204,
                        69,
                        156)),

            "extract_pdf_pages" =>
                BuildVectorIcon(
                    GetPagesGeometry(),
                    WpfColor.FromRgb(
                        83,
                        118,
                        163)),

            "merge_pdf" =>
                BuildVectorIcon(
                    GetLinkGeometry(),
                    WpfColor.FromRgb(
                        84,
                        70,
                        211)),

            "extract_zip" =>
                BuildVectorIcon(
                    GetExtractArchiveGeometry(),
                    WpfColor.FromRgb(
                        38,
                        132,
                        156)),

            "zip" =>
                BuildVectorIcon(
                    GetArchiveGeometry(),
                    WpfColor.FromRgb(
                        84,
                        99,
                        121)),


            "pdf_to_txt" =>
                BuildOfficeIcon(
                    "T",
                    WpfColor.FromRgb(
                        102,
                        116,
                        136)),

            "to_png" =>
                BuildVectorIcon(
                    GetImageGeometry(),
                    WpfColor.FromRgb(
                        104,
                        93,
                        215)),

            "to_jpg" =>
                BuildVectorIcon(
                    GetImageGeometry(),
                    WpfColor.FromRgb(
                        184,
                        91,
                        138)),

            "to_webp" =>
                BuildOfficeIcon(
                    "W",
                    WpfColor.FromRgb(
                        43,
                        143,
                        140)),

            "to_pdf" =>
                BuildOfficeIcon(
                    "PDF",
                    WpfColor.FromRgb(
                        215,
                        67,
                        82),
                    7),

            "edit_image" =>
                BuildVectorIcon(
                    GetEditGeometry(),
                    WpfColor.FromRgb(
                        55,
                        136,
                        168)),

            _ =>
                BuildOfficeIcon(
                    "•",
                    WpfColor.FromRgb(
                        100,
                        115,
                        135))
        };
    }


    private FrameworkElement BuildOfficeIcon(
        string text,
        WpfColor color,
        double fontSize = 15)
    {
        var root =
            new WpfGrid
            {
                Width =
                    39,

                Height =
                    39,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center
            };

        var ghost =
            new WpfBorder
            {
                Width =
                    27,

                Height =
                    27,

                CornerRadius =
                    new CornerRadius(
                        8),

                Background =
                    new SolidColorBrush(
                        WpfColor.FromArgb(
                            72,
                            color.R,
                            color.G,
                            color.B)),

                HorizontalAlignment =
                    WpfHorizontalAlignment.Right,

                VerticalAlignment =
                    WpfVerticalAlignment.Top
            };

        var tile =
            new WpfBorder
            {
                Width =
                    34,

                Height =
                    34,

                CornerRadius =
                    new CornerRadius(
                        9),

                Background =
                    CreateIconGradient(
                        color),

                HorizontalAlignment =
                    WpfHorizontalAlignment.Left,

                VerticalAlignment =
                    WpfVerticalAlignment.Bottom
            };

        tile.Effect =
            new System.Windows.Media.Effects.DropShadowEffect
            {
                Color =
                    color,

                BlurRadius =
                    7,

                ShadowDepth =
                    2,

                Opacity =
                    0.17
            };

        tile.Child =
            new WpfTextBlock
            {
                Text =
                    text,

                FontFamily =
                    new WpfFontFamily(
                        "Segoe UI"),

                FontSize =
                    fontSize,

                FontWeight =
                    FontWeights.SemiBold,

                Foreground =
                    WpfBrushes.White,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                VerticalAlignment =
                    WpfVerticalAlignment.Center,

                TextAlignment =
                    TextAlignment.Center
            };

        root.Children.Add(
            ghost);

        root.Children.Add(
            tile);

        return root;
    }


    private FrameworkElement BuildVectorIcon(
        WpfGeometry geometry,
        WpfColor color)
    {
        var root =
            new WpfGrid
            {
                Width =
                    39,

                Height =
                    39,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center
            };

        var halo =
            new WpfBorder
            {
                Width =
                    37,

                Height =
                    37,

                CornerRadius =
                    new CornerRadius(
                        12),

                Background =
                    new SolidColorBrush(
                        WpfColor.FromArgb(
                            24,
                            color.R,
                            color.G,
                            color.B)),

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                VerticalAlignment =
                    WpfVerticalAlignment.Center
            };

        var tile =
            new WpfBorder
            {
                Width =
                    33,

                Height =
                    33,

                CornerRadius =
                    new CornerRadius(
                        9),

                Background =
                    CreateIconGradient(
                        color),

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                VerticalAlignment =
                    WpfVerticalAlignment.Center
            };

        tile.Effect =
            new System.Windows.Media.Effects.DropShadowEffect
            {
                Color =
                    color,

                BlurRadius =
                    7,

                ShadowDepth =
                    2,

                Opacity =
                    0.16
            };

        var path =
            new WpfPath
            {
                Data =
                    geometry,

                Stroke =
                    WpfBrushes.White,

                StrokeThickness =
                    1.8,

                StrokeStartLineCap =
                    PenLineCap.Round,

                StrokeEndLineCap =
                    PenLineCap.Round,

                StrokeLineJoin =
                    PenLineJoin.Round,

                Fill =
                    WpfBrushes.Transparent,

                Width =
                    19,

                Height =
                    19,

                Stretch =
                    Stretch.Uniform,

                HorizontalAlignment =
                    WpfHorizontalAlignment.Center,

                VerticalAlignment =
                    WpfVerticalAlignment.Center
            };

        tile.Child =
            path;

        root.Children.Add(
            halo);

        root.Children.Add(
            tile);

        return root;
    }


    private static WpfBrush CreateIconGradient(
        WpfColor color)
    {
        var lighter =
            Lighten(
                color,
                35);

        return new WpfLinearGradientBrush
        {
            StartPoint =
                new WpfPoint(
                    0,
                    0),

            EndPoint =
                new WpfPoint(
                    1,
                    1),

            GradientStops =
                new WpfGradientStopCollection
                {
                    new WpfGradientStop(
                        lighter,
                        0),

                    new WpfGradientStop(
                        color,
                        1)
                }
        };
    }


    private static WpfColor Lighten(
        WpfColor color,
        byte amount)
    {
        return WpfColor.FromRgb(
            (byte)Math.Min(
                255,
                color.R + amount),

            (byte)Math.Min(
                255,
                color.G + amount),

            (byte)Math.Min(
                255,
                color.B + amount));
    }


    // ====================================================
    // VECTOR GEOMETRY
    // ====================================================

    private static WpfGeometry GetImageGeometry()
    {
        return WpfGeometry.Parse(
            "M3,4 L21,4 L21,20 L3,20 Z " +
            "M6,16 L10,12 L13,15 L16,11 L20,16 " +
            "M8,8 A1.5,1.5 0 1 1 7.99,8");
    }


    private static WpfGeometry GetScanGeometry()
    {
        return WpfGeometry.Parse(
            "M3,8 L3,3 L8,3 " +
            "M16,3 L21,3 L21,8 " +
            "M21,16 L21,21 L16,21 " +
            "M8,21 L3,21 L3,16 " +
            "M7,12 L17,12");
    }


    private static WpfGeometry GetScissorsGeometry()
    {
        return WpfGeometry.Parse(
            "M6,5 A3,3 0 1 1 5.99,5 " +
            "M6,19 A3,3 0 1 1 5.99,19 " +
            "M8.5,7 L20,18 " +
            "M8.5,17 L20,6");
    }


    private static WpfGeometry GetRotateGeometry()
    {
        return WpfGeometry.Parse(
            "M19,8 L19,3 L23,7 L19,11 L19,8 " +
            "M19,7 C17,4 14,3 11,3 C6,3 2,7 2,12 C2,17 6,21 11,21 C15,21 18,19 20,16");
    }


    private static WpfGeometry GetTrashGeometry()
    {
        return WpfGeometry.Parse(
            "M5,7 L19,7 " +
            "M9,7 L9,4 L15,4 L15,7 " +
            "M7,7 L8,21 L16,21 L17,7 " +
            "M10,11 L10.5,18 " +
            "M14,11 L13.5,18");
    }


    private static WpfGeometry GetReorderGeometry()
    {
        return WpfGeometry.Parse(
            "M4,5 A1,1 0 1 1 3.99,5 " +
            "M8,5 L21,5 " +
            "M4,12 A1,1 0 1 1 3.99,12 " +
            "M8,12 L21,12 " +
            "M4,19 A1,1 0 1 1 3.99,19 " +
            "M8,19 L21,19");
    }


    private static WpfGeometry GetPagesGeometry()
    {
        return WpfGeometry.Parse(
            "M7,4 L19,4 L19,18 L7,18 Z " +
            "M4,7 L4,21 L16,21 " +
            "M10,8 L16,8 " +
            "M10,12 L16,12");
    }


    private static WpfGeometry GetLinkGeometry()
    {
        return WpfGeometry.Parse(
            "M9,15 L7,17 A4,4 0 0 1 1,11 L5,7 A4,4 0 0 1 11,7 " +
            "M15,9 L17,7 A4,4 0 0 1 23,13 L19,17 A4,4 0 0 1 13,17 " +
            "M8,12 L16,12");
    }


    private static WpfGeometry GetExtractArchiveGeometry()
    {
        return WpfGeometry.Parse(
            "M5,3 L16,3 L20,7 L20,13 " +
            "M16,3 L16,7 L20,7 " +
            "M9,8 L12,8 M9,11 L12,11 " +
            "M12,13 L12,21 " +
            "M8.5,17.5 L12,21 L15.5,17.5 " +
            "M5,21 L8,21 M16,21 L20,21");
    }

    private static WpfGeometry GetArchiveGeometry()
    {
        return WpfGeometry.Parse(
            "M6,2 L15,2 L20,7 L20,22 L6,22 Z " +
            "M15,2 L15,7 L20,7 " +
            "M11,5 L14,5 " +
            "M11,8 L14,8 " +
            "M11,11 L14,11 " +
            "M11,14 L14,14 " +
            "M10,18 L15,18");
    }


    private static WpfGeometry GetEditGeometry()
    {
        return WpfGeometry.Parse(
            "M5,19 L8,18 L19,7 L16,4 L5,15 Z " +
            "M14.5,5.5 L17.5,8.5 " +
            "M5,19 L9,18");
    }


    // ====================================================
    // LABELS
    // ====================================================

    private static string GetActionTitle(
        FileAction action)
    {
        return action.Id switch
        {
            "pdf_to_docx" =>
                "To DOCX",

            "pdf_to_ocr_docx" =>
                "OCR to DOCX",

            "pdf_to_xlsx" =>
                "To XLSX",

            "pdf_to_pptx" =>
                "To PPTX",

            "pdf_to_txt" =>
                "To TXT",

            "pdf_to_png" =>
                "To PNG",

            "pdf_to_jpg" =>
                "To JPG",

            "extract_pdf_images" =>
                "Extract Images",

            "split_pdf" =>
                "Split Pages",

            "rotate_pdf_pages" =>
                "Rotate Pages",

            "delete_pdf_pages" =>
                "Delete Pages",

            "reorder_pdf_pages" =>
                "Reorder Pages",

            "extract_pdf_pages" =>
                "Extract Pages",

            "merge_pdf" =>
                "Merge PDFs",

            "extract_zip" =>
                "Extract ZIP",

            "zip" =>
                "ZIP Files",

            "to_png" =>
                "To PNG",

            "to_jpg" =>
                "To JPG",

            "to_webp" =>
                "To WEBP",

            "to_pdf" =>
                "To PDF",

            "edit_image" =>
                "Edit Image",

            _ =>
                action.DisplayName
        };
    }


    private static string GetActionSubtitle(
        FileAction action)
    {
        return action.Id switch
        {
            "pdf_to_docx" =>
                "Text extraction",

            "pdf_to_ocr_docx" =>
                "Scanned PDF",

            "pdf_to_xlsx" =>
                "Tables to Excel",

            "pdf_to_pptx" =>
                "Editable",

            "pdf_to_txt" =>
                "Plain text",

            "pdf_to_png" =>
                "All pages",

            "pdf_to_jpg" =>
                "All pages",

            "extract_pdf_images" =>
                "From PDF",

            "split_pdf" =>
                "Every page",

            "rotate_pdf_pages" =>
                "Custom angle",

            "delete_pdf_pages" =>
                "By range",

            "reorder_pdf_pages" =>
                "Drag & drop",

            "extract_pdf_pages" =>
                "By range",

            "merge_pdf" =>
                "Combine files",

            "extract_zip" =>
                "Unpack files",

            "zip" =>
                "Compress",

            "to_png" =>
                "Convert image",

            "to_jpg" =>
                "Convert image",

            "to_webp" =>
                "Convert image",

            "to_pdf" =>
                "Create PDF",

            "edit_image" =>
                "Crop & adjust",

            _ =>
                string.Empty
        };
    }


    // ====================================================
    // COLORS
    // ====================================================

    private static WpfBrush GetActionAccentBrush(
        FileAction action)
    {
        var color =
            action.Id switch
            {
                "pdf_to_docx" =>
                    WpfColor.FromRgb(
                        36,
                        103,
                        190),

                "pdf_to_ocr_docx" =>
                    WpfColor.FromRgb(
                        45,
                        143,
                        141),

                "pdf_to_xlsx" =>
                    WpfColor.FromRgb(
                        33,
                        129,
                        92),

                "pdf_to_pptx" =>
                    WpfColor.FromRgb(
                        210,
                        72,
                        54),

                "pdf_to_txt" =>
                    WpfColor.FromRgb(
                        102,
                        116,
                        136),

                "pdf_to_png" =>
                    WpfColor.FromRgb(
                        104,
                        93,
                        215),

                "pdf_to_jpg" =>
                    WpfColor.FromRgb(
                        184,
                        91,
                        138),

                "extract_pdf_images" =>
                    WpfColor.FromRgb(
                        214,
                        157,
                        59),

                "split_pdf" =>
                    WpfColor.FromRgb(
                        116,
                        82,
                        206),

                "rotate_pdf_pages" =>
                    WpfColor.FromRgb(
                        56,
                        125,
                        212),

                "delete_pdf_pages" =>
                    WpfColor.FromRgb(
                        215,
                        67,
                        82),

                "reorder_pdf_pages" =>
                    WpfColor.FromRgb(
                        204,
                        69,
                        156),

                "extract_pdf_pages" =>
                    WpfColor.FromRgb(
                        83,
                        118,
                        163),

                "merge_pdf" =>
                    WpfColor.FromRgb(
                        84,
                        70,
                        211),

                "zip" =>
                    WpfColor.FromRgb(
                        84,
                        99,
                        121),

                _ =>
                    WpfColor.FromRgb(
                        100,
                        115,
                        135)
            };

        return new SolidColorBrush(
            color);
    }


    private static SolidColorBrush
        CreateNormalButtonBrush()
    {
        return new SolidColorBrush(
            WpfColor.FromArgb(
                250,
                232,
                238,
                245));
    }


    private static SolidColorBrush
        CreateNormalBorderBrush()
    {
        return new SolidColorBrush(
            WpfColor.FromArgb(
                235,
                164,
                179,
                198));
    }


    // ====================================================
    // OPEN ANIMATION
    // ====================================================

    private void OnOverlayVisibilityChanged(
        object sender,
        DependencyPropertyChangedEventArgs e)
    {
        if (e.NewValue
            is not bool visible)
        {
            return;
        }

        //
        // Анимацию появления здесь больше
        // НЕ запускаем.
        //
        // Она запускается явно только после
        // правильного позиционирования окна.
        //
        if (!visible)
        {
            DraggedFiles =
                Array.Empty<string>();

            _activeActionButton =
                null;

            StopOpenAnimation();
        }
    }


    private void PrepareOpenAnimationState()
    {
        //
        // Полностью останавливаем предыдущие
        // анимации, если они вдруг остались.
        //
        RootContainer.BeginAnimation(
            OpacityProperty,
            null);

        var scale =
            new WpfScaleTransform(
                0.97,
                0.97);

        RootContainer.RenderTransform =
            scale;

        RootContainer.RenderTransformOrigin =
            new WpfPoint(
                0.5,
                0.5);

        //
        // Важно: opacity выставляется ДО Show().
        //
        RootContainer.Opacity =
            0;
    }


    private void PlayOpenAnimation()
    {
        var scale =
            RootContainer.RenderTransform
                as WpfScaleTransform;

        if (scale is null)
        {
            scale =
                new WpfScaleTransform(
                    0.97,
                    0.97);

            RootContainer.RenderTransform =
                scale;
        }

        RootContainer.RenderTransformOrigin =
            new WpfPoint(
                0.5,
                0.5);

        var easing =
            new CubicEase
            {
                EasingMode =
                    EasingMode.EaseOut
            };

        var scaleX =
            new DoubleAnimation
            {
                From =
                    scale.ScaleX,

                To =
                    1.0,

                Duration =
                    TimeSpan.FromMilliseconds(
                        125),

                EasingFunction =
                    easing
            };

        var scaleY =
            new DoubleAnimation
            {
                From =
                    scale.ScaleY,

                To =
                    1.0,

                Duration =
                    TimeSpan.FromMilliseconds(
                        125),

                EasingFunction =
                    easing
            };

        var opacity =
            new DoubleAnimation
            {
                From =
                    RootContainer.Opacity,

                To =
                    1,

                Duration =
                    TimeSpan.FromMilliseconds(
                        95),

                EasingFunction =
                    easing
            };

        scale.BeginAnimation(
            WpfScaleTransform.ScaleXProperty,
            scaleX);

        scale.BeginAnimation(
            WpfScaleTransform.ScaleYProperty,
            scaleY);

        RootContainer.BeginAnimation(
            OpacityProperty,
            opacity);
    }


    private void StopOpenAnimation()
    {
        RootContainer.BeginAnimation(
            OpacityProperty,
            null);

        if (RootContainer.RenderTransform
            is WpfScaleTransform scale)
        {
            scale.BeginAnimation(
                WpfScaleTransform.ScaleXProperty,
                null);

            scale.BeginAnimation(
                WpfScaleTransform.ScaleYProperty,
                null);

            scale.ScaleX =
                1;

            scale.ScaleY =
                1;
        }

        RootContainer.Opacity =
            1;
    }

    // ====================================================
    // DRAG
    // ====================================================

    private void OnPreviewDragEnter(
        object sender,
        System.Windows.DragEventArgs e)
    {
        ReadDraggedFiles(
            e);

        if (DraggedFiles.Count == 0)
        {
            SetActiveButton(
                null);

            return;
        }

        UpdateActiveButtonFromDrag(
            e);
    }


    private void OnPreviewDragOver(
        object sender,
        System.Windows.DragEventArgs e)
    {
        if (DraggedFiles.Count == 0)
        {
            ReadDraggedFiles(
                e);
        }

        if (DraggedFiles.Count == 0)
        {
            SetActiveButton(
                null);

            e.Effects =
                WpfDragDropEffects.None;

            e.Handled =
                true;

            return;
        }

        UpdateActiveButtonFromDrag(
            e);

        e.Effects =
            WpfDragDropEffects.Copy;

        e.Handled =
            true;
    }


    private void UpdateActiveButtonFromDrag(
        System.Windows.DragEventArgs e)
    {
        var button =
            FindParentButton(
                e.OriginalSource
                    as DependencyObject);

        if (button?.Tag
            is FileAction)
        {
            SetActiveButton(
                button);
        }
        else
        {
            SetActiveButton(
                null);
        }
    }


    // ====================================================
    // DROP
    // ====================================================


    private async void OnPreviewDrop(
        object sender,
        System.Windows.DragEventArgs e)
    {
        if (DraggedFiles.Count == 0)
        {
            ReadDraggedFiles(e);
        }

        if (DraggedFiles.Count == 0)
        {
            SetActiveButton(null);
            e.Effects = WpfDragDropEffects.None;
            e.Handled = true;
            return;
        }

        var button = FindParentButton(
            e.OriginalSource as DependencyObject);

        if (button?.Tag is not FileAction action)
        {
            SetActiveButton(null);
            e.Effects = WpfDragDropEffects.None;
            e.Handled = true;
            return;
        }

        e.Effects = WpfDragDropEffects.Copy;
        e.Handled = true;

        // Сохраняем информацию до выполнения действия.
        var sourceFiles = DraggedFiles.ToArray();

        var fileDescription =
            sourceFiles.Length == 1
                ? Path.GetFileName(sourceFiles[0])
                : $"{sourceFiles.Length} files";

        SetActiveButton(null);

        try
        {
            // Метод возвращает false при отмене.
            bool completed =
                await ExecuteActionAsync(action);

            if (!completed)
                return;

            string title = action.Id switch
            {
                "zip" => "Archive created",
                "extract_zip" => "Extraction completed",
                "merge_pdf" => "PDFs merged",
                "split_pdf" => "PDF split completed",
                "edit_image" => "Image saved",
                _ => "Conversion completed"
            };

            NotificationRequested?.Invoke(
                title,
                $"{fileDescription} processed successfully.",
                false);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine(
                $"Action failed: {ex}");

            NotificationRequested?.Invoke(
                "Operation failed",
                $"{fileDescription}: {ex.Message}",
                true);

            WpfMessageBox.Show(
                ex.Message,
                "Morph",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
    }



    // ====================================================
    // FILE READING
    // ====================================================

    private void ReadDraggedFiles(
        System.Windows.DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(
                System.Windows.DataFormats.FileDrop))
        {
            if (DraggedFiles.Count == 0)
            {
                e.Effects =
                    WpfDragDropEffects.None;
            }

            e.Handled =
                true;

            return;
        }

        var data =
            e.Data.GetData(
                System.Windows.DataFormats.FileDrop);

        if (data
            is not string[] paths)
        {
            e.Handled =
                true;

            return;
        }

        var files =
            paths
                .Where(
                    File.Exists)
                .ToArray();

        if (files.Length == 0)
        {
            e.Handled =
                true;

            return;
        }

        var filesChanged =
            DraggedFiles.Count !=
            files.Length ||
            !DraggedFiles.SequenceEqual(
                files,
                StringComparer.OrdinalIgnoreCase);

        if (filesChanged)
        {
            DraggedFiles =
                files;

            UpdateFileInfo(
                files);

            UpdateActions(
                files);

            DragFilesDetected?.Invoke(
                files);
        }

        e.Effects =
            WpfDragDropEffects.Copy;

        e.Handled =
            true;
    }


    // ====================================================
    // FILE INFO
    // ====================================================

    private void UpdateFileInfo(
        IReadOnlyList<string> files)
    {
        if (files.Count == 0)
        {
            FileTypeText.Text =
                "DROP";

            FileNameText.Text =
                "Drag file";

            FileMetaText.Text =
                "drop a file to begin";

            FileTypeBadge.Background =
                new SolidColorBrush(
                    WpfColor.FromRgb(
                        105,
                        117,
                        134));

            return;
        }

        if (files.Count == 1)
        {
            var path =
                files[0];

            var extension =
                Path.GetExtension(
                    path)
                .TrimStart('.')
                .ToUpperInvariant();

            if (string.IsNullOrWhiteSpace(
                    extension))
            {
                extension =
                    "FILE";
            }

            FileTypeText.Text =
                extension;

            FileNameText.Text =
                Path.GetFileName(
                    path);

            var info =
                new FileInfo(
                    path);

            FileMetaText.Text =
                $"1 file  •  {FormatFileSize(info.Length)}";

            FileTypeBadge.Background =
                GetFileTypeBadgeBrush(
                    extension);

            return;
        }

        long totalSize =
            0;

        foreach (var file in files)
        {
            try
            {
                totalSize +=
                    new FileInfo(
                        file)
                    .Length;
            }
            catch
            {
            }
        }

        FileTypeText.Text =
            files.Count.ToString();

        FileNameText.Text =
            $"{files.Count} files";

        FileMetaText.Text =
            FormatFileSize(
                totalSize);

        FileTypeBadge.Background =
            new SolidColorBrush(
                WpfColor.FromRgb(
                    95,
                    108,
                    126));
    }


    private static WpfBrush
        GetFileTypeBadgeBrush(
            string extension)
    {
        var color =
            extension switch
            {
                "PDF" =>
                    WpfColor.FromRgb(
                        215,
                        76,
                        80),

                "DOCX" =>
                    WpfColor.FromRgb(
                        36,
                        103,
                        190),

                "XLSX" =>
                    WpfColor.FromRgb(
                        33,
                        129,
                        92),

                "PPTX" =>
                    WpfColor.FromRgb(
                        210,
                        72,
                        54),

                "PNG" =>
                    WpfColor.FromRgb(
                        104,
                        93,
                        215),

                "JPG" or "JPEG" =>
                    WpfColor.FromRgb(
                        184,
                        91,
                        138),

                "WEBP" =>
                    WpfColor.FromRgb(
                        45,
                        143,
                        141),

                _ =>
                    WpfColor.FromRgb(
                        105,
                        117,
                        134)
            };

        return new SolidColorBrush(
            color);
    }


    private static string FormatFileSize(
        long bytes)
    {
        if (bytes <
            1024)
        {
            return
                $"{bytes} B";
        }

        if (bytes <
            1024L *
            1024L)
        {
            return
                $"{bytes / 1024.0:0.#} KB";
        }

        if (bytes <
            1024L *
            1024L *
            1024L)
        {
            return
                $"{bytes / 1024.0 / 1024.0:0.#} MB";
        }

        return
            $"{bytes / 1024.0 / 1024.0 / 1024.0:0.##} GB";
    }


    // ====================================================
    // ACTIONS
    // ====================================================

    private void UpdateActions(
        IReadOnlyList<string> files)
    {
        var actions =
            _actionRegistry.GetActions(
                files);

        var buttons =
            GetActionButtons();

        _activeActionButton =
            null;

        foreach (var button in
                 buttons)
        {
            button.BeginAnimation(
                OpacityProperty,
                null);

            button.Opacity =
                1;

            button.Content =
                null;

            button.Tag =
                null;

            button.Visibility =
                Visibility.Collapsed;

            button.Background =
                CreateNormalButtonBrush();

            button.BorderBrush =
                CreateNormalBorderBrush();

            button.BorderThickness =
                new Thickness(
                    1);

            button.RenderTransform =
                new WpfScaleTransform(
                    1,
                    1);

            button.RenderTransformOrigin =
                new WpfPoint(
                    0.5,
                    0.5);
        }

        for (var i = 0;
             i < actions.Count &&
             i < buttons.Length;
             i++)
        {
            var button =
                buttons[i];

            button.Tag =
                actions[i];

            button.Visibility =
                Visibility.Visible;
        }

        RefreshActionButtonContents();

        ArrangeActionButtons();
    }


    // ====================================================
    // ACTIVE BUTTON
    // ====================================================

    private void SetActiveButton(
        WpfButton? button)
    {
        if (_activeActionButton ==
            button)
        {
            return;
        }

        var previous =
            _activeActionButton;

        _activeActionButton =
            button;

        foreach (var actionButton in
                 GetActionButtons())
        {
            if (actionButton.Visibility !=
                Visibility.Visible)
            {
                continue;
            }

            if (actionButton != button &&
                actionButton != previous)
            {
                actionButton.BeginAnimation(
                    OpacityProperty,
                    null);

                actionButton.Opacity =
                    button is null
                        ? 1.0
                        : 0.72;

                EnsureButtonScale(
                    actionButton,
                    1.0);
            }
        }

        if (previous is not null &&
            previous.Visibility ==
            Visibility.Visible)
        {
            AnimateButton(
                previous,
                false,
                button is not null
                    ? 0.72
                    : 1.0);
        }

        if (button is not null)
        {
            AnimateButton(
                button,
                true,
                1.0);
        }

        if (button is null)
        {
            foreach (var actionButton in
                     GetActionButtons())
            {
                if (actionButton.Visibility !=
                    Visibility.Visible)
                {
                    continue;
                }

                actionButton.Opacity =
                    1.0;

                actionButton.Background =
                    CreateNormalButtonBrush();

                actionButton.BorderBrush =
                    CreateNormalBorderBrush();

                actionButton.BorderThickness =
                    new Thickness(
                        1);

                EnsureButtonScale(
                    actionButton,
                    1.0);
            }
        }
    }


    private void AnimateButton(
        WpfButton button,
        bool active,
        double targetOpacity)
    {
        var scale =
            GetMutableScale(
                button);

        var easing =
            new CubicEase
            {
                EasingMode =
                    EasingMode.EaseOut
            };

        var targetScale =
            active
                ? 1.075
                : 1.0;

        var scaleX =
            new DoubleAnimation
            {
                To =
                    targetScale,

                Duration =
                    TimeSpan.FromMilliseconds(
                        110),

                EasingFunction =
                    easing
            };

        var scaleY =
            new DoubleAnimation
            {
                To =
                    targetScale,

                Duration =
                    TimeSpan.FromMilliseconds(
                        110),

                EasingFunction =
                    easing
            };

        var opacity =
            new DoubleAnimation
            {
                To =
                    targetOpacity,

                Duration =
                    TimeSpan.FromMilliseconds(
                        90),

                EasingFunction =
                    easing
            };

        scale.BeginAnimation(
            WpfScaleTransform.ScaleXProperty,
            scaleX);

        scale.BeginAnimation(
            WpfScaleTransform.ScaleYProperty,
            scaleY);

        button.BeginAnimation(
            OpacityProperty,
            opacity);

        if (button.Tag
            is not FileAction action)
        {
            return;
        }

        if (active)
        {
            button.Background =
                new SolidColorBrush(
                    WpfColor.FromArgb(
                        252,
                        250,
                        251,
                        253));

            button.BorderBrush =
                GetActionAccentBrush(
                    action);

            button.BorderThickness =
                new Thickness(
                    1.6);
        }
        else
        {
            button.Background =
                CreateNormalButtonBrush();

            button.BorderBrush =
                CreateNormalBorderBrush();

            button.BorderThickness =
                new Thickness(
                    1);
        }
    }


    private static WpfScaleTransform
        GetMutableScale(
            WpfButton button)
    {
        if (button.RenderTransform
                is WpfScaleTransform existing &&
            !existing.IsFrozen)
        {
            return existing;
        }

        var scale =
            new WpfScaleTransform(
                1,
                1);

        button.RenderTransform =
            scale;

        button.RenderTransformOrigin =
            new WpfPoint(
                0.5,
                0.5);

        return scale;
    }


    private static void EnsureButtonScale(
        WpfButton button,
        double value)
    {
        var scale =
            GetMutableScale(
                button);

        scale.BeginAnimation(
            WpfScaleTransform.ScaleXProperty,
            null);

        scale.BeginAnimation(
            WpfScaleTransform.ScaleYProperty,
            null);

        scale.ScaleX =
            value;

        scale.ScaleY =
            value;
    }


    // ====================================================
    // HIT TEST
    // ====================================================

    private static WpfButton? FindParentButton(
        DependencyObject? element)
    {
        while (element is not null)
        {
            if (element
                is WpfButton button)
            {
                return button;
            }

            element =
                WpfVisualTreeHelper.GetParent(
                    element);
        }

        return null;
    }


    private void OnActionClick(
        object sender,
        RoutedEventArgs e)
    {
        //
        // Основное действие выполняется Drop'ом.
        //
    }


    // ====================================================
    // EXECUTION
    // ====================================================


    private async Task<bool> ExecuteActionAsync(FileAction action)
    {
        var files = DraggedFiles.ToArray();

        if (files.Length == 0)
            return false;

        switch (action.Id)
        {
            case "to_png":
            case "to_jpg":
            case "to_webp":
                {
                    if (files.Length != 1)
                        return false;

                    var outputPath =
                        await _imageConversionService.ConvertAsync(
                            files[0], action.Id);

                    System.Diagnostics.Debug.WriteLine(
                        $"Converted: {outputPath}");

                    break;
                }

            case "pdf_to_png":
            case "pdf_to_jpg":
                {
                    if (files.Length != 1)
                        return false;

                    var outputDirectory =
                        await _pdfRenderService.RenderToImagesAsync(
                            files[0], action.Id);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF pages rendered: {outputDirectory}");

                    break;
                }

            case "to_pdf":
                {
                    if (files.Length != 1)
                        return false;

                    var sourcePath = files[0];
                    var extension = Path.GetExtension(sourcePath)
                        .ToLowerInvariant();

                    string outputPath;

                    if (extension is ".jpg" or ".jpeg" or ".png" or ".webp")
                    {
                        outputPath =
                            await _pdfConversionService
                                .ConvertImageToPdfAsync(sourcePath);
                    }
                    else if (extension is ".docx" or ".xlsx" or ".pptx")
                    {
                        outputPath =
                            await _officeConversionService
                                .ConvertToPdfAsync(sourcePath);
                    }
                    else
                    {
                        throw new NotSupportedException(
                            $"PDF conversion is not supported for {extension}.");
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF created: {outputPath}");

                    break;
                }

            case "pdf_to_txt":
                {
                    if (files.Length != 1)
                        return false;

                    var outputPath =
                        await _pdfExtractionService
                            .ExtractToTxtAsync(files[0]);

                    System.Diagnostics.Debug.WriteLine(
                        $"TXT extracted: {outputPath}");

                    break;
                }

            case "pdf_to_docx":
                {
                    if (files.Length != 1)
                        return false;

                    var outputPath =
                        await _pdfExtractionService
                            .ExtractToDocxAsync(files[0]);

                    System.Diagnostics.Debug.WriteLine(
                        $"DOCX extracted: {outputPath}");

                    break;
                }

            case "merge_pdf":
                {
                    var outputPath =
                        await _pdfMergeService.MergeAsync(files);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF merged: {outputPath}");

                    break;
                }

            case "extract_pdf_images":
                {
                    if (files.Length != 1)
                        return false;

                    var outputDirectory =
                        await _pdfImageExtractionService
                            .ExtractAsync(files[0]);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF images extracted: {outputDirectory}");

                    break;
                }

            case "split_pdf":
                {
                    if (files.Length != 1)
                        return false;

                    var outputDirectory =
                        await _pdfSplitService.SplitAsync(files[0]);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF split: {outputDirectory}");

                    break;
                }

            case "pdf_to_ocr_docx":
                {
                    if (files.Length != 1)
                        return false;

                    var outputPath =
                        await _pdfOcrDocxService
                            .ConvertAsync(files[0]);

                    System.Diagnostics.Debug.WriteLine(
                        $"OCR DOCX created: {outputPath}");

                    break;
                }

            case "reorder_pdf_pages":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new ReorderPagesDialog();

                    if (dialog.ShowDialog() != true)
                        return false;

                    var pages =
                        _pdfPageExtractionService
                            .ParsePageRange(dialog.PageOrder);

                    var outputPath =
                        await _pdfPageReorderService
                            .ReorderAsync(files[0], pages);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF pages reordered: {outputPath}");

                    break;
                }

            case "rotate_pdf_pages":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new RotatePagesDialog();

                    if (dialog.ShowDialog() != true)
                        return false;

                    var pages =
                        _pdfPageExtractionService
                            .ParsePageRange(dialog.PageRange);

                    var outputPath =
                        await _pdfPageRotationService.RotateAsync(
                            files[0],
                            pages,
                            dialog.RotationDegrees);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF pages rotated: {outputPath}");

                    break;
                }

            case "extract_pdf_pages":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new PageRangeDialog();

                    if (dialog.ShowDialog() != true)
                        return false;

                    var pages =
                        _pdfPageExtractionService
                            .ParsePageRange(dialog.PageRange);

                    var outputPath =
                        await _pdfPageExtractionService
                            .ExtractAsync(files[0], pages);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF pages extracted: {outputPath}");

                    break;
                }

            case "pdf_to_xlsx":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new PdfToExcelDialog();

                    if (dialog.ShowDialog() != true)
                        return false;

                    var outputPath =
                        await _pdfExcelExtractionService
                            .ConvertAsync(files[0], dialog.ExportMode);

                    System.Diagnostics.Debug.WriteLine(
                        $"XLSX created: {outputPath}");

                    break;
                }

            case "delete_pdf_pages":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new DeletePagesDialog();

                    if (dialog.ShowDialog() != true)
                        return false;

                    var pages =
                        _pdfPageExtractionService
                            .ParsePageRange(dialog.PageRange);

                    var outputPath =
                        await _pdfPageDeletionService
                            .DeleteAsync(files[0], pages);

                    System.Diagnostics.Debug.WriteLine(
                        $"PDF pages deleted: {outputPath}");

                    break;
                }

            case "pdf_to_pptx":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new PdfToPowerPointDialog();

                    if (dialog.ShowDialog() != true)
                        return false;

                    if (dialog.ExportMode == PdfPowerPointExportMode.Exact)
                    {
                        var outputPath =
                            await _pdfPowerPointService
                                .ConvertAsync(files[0]);

                        System.Diagnostics.Debug.WriteLine(
                            $"PPTX created: {outputPath}");
                    }
                    else
                    {
                        var outputPath =
                            await _pdfEditablePowerPointService
                                .ConvertAsync(files[0]);

                        System.Diagnostics.Debug.WriteLine(
                            $"Editable PPTX created: {outputPath}");
                    }

                    break;
                }

            case "edit_image":
                {
                    if (files.Length != 1)
                        return false;

                    var dialog = new ImageEditorDialog(files[0]);

                    if (dialog.ShowDialog() != true)
                        return false;

                    if (string.IsNullOrWhiteSpace(dialog.OutputPath))
                        return false;

                    System.Diagnostics.Debug.WriteLine(
                        $"Image edited: {dialog.OutputPath}");

                    break;
                }

            case "extract_zip":
                {
                    var outputDirectories =
                        await _zipService.ExtractAsync(files);

                    foreach (var directory in outputDirectories)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"ZIP extracted: {directory}");
                    }

                    break;
                }

            case "zip":
                {
                    var outputPath =
                        await _zipService.CreateAsync(files);

                    System.Diagnostics.Debug.WriteLine(
                        $"ZIP created: {outputPath}");

                    break;
                }

            default:
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"Action not implemented yet: {action.Id}");

                    return false;
                }
        }

        return true;
    }



    // ====================================================
    // RESET
    // ====================================================

    private void ClearDraggedFiles()
    {
        DraggedFiles =
            Array.Empty<string>();

        ResetUi();
    }


    private void ResetUi()
    {
        _activeActionButton =
            null;

        ConnectorCanvas.Children.Clear();

        FileTypeText.Text =
            "DROP";

        FileNameText.Text =
            "Drag file";

        FileMetaText.Text =
            "drop a file to begin";

        FileTypeBadge.Background =
            new SolidColorBrush(
                WpfColor.FromRgb(
                    105,
                    117,
                    134));

        foreach (var button in
                 GetActionButtons())
        {
            button.BeginAnimation(
                OpacityProperty,
                null);

            button.Opacity =
                1;

            button.Visibility =
                Visibility.Collapsed;

            button.Content =
                null;

            button.Tag =
                null;

            button.Background =
                CreateNormalButtonBrush();

            button.BorderBrush =
                CreateNormalBorderBrush();

            button.BorderThickness =
                new Thickness(
                    1);

            button.RenderTransform =
                new WpfScaleTransform(
                    1,
                    1);

            button.RenderTransformOrigin =
                new WpfPoint(
                    0.5,
                    0.5);
        }
    }


    // ====================================================
    // NATIVE WINDOW POSITIONING
    // ====================================================

    private const uint SWP_NOSIZE =
        0x0001;

    private const uint SWP_NOZORDER =
        0x0004;

    private const uint SWP_NOACTIVATE =
        0x0010;


    [StructLayout(
        LayoutKind.Sequential)]
    private struct NativeWindowRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }


    [DllImport(
        "user32.dll",
        SetLastError = true)]
    private static extern bool GetWindowRect(
        IntPtr hWnd,
        out NativeWindowRect lpRect);


    [DllImport(
        "user32.dll",
        SetLastError = true)]
    private static extern bool SetWindowPos(
        IntPtr hWnd,
        IntPtr hWndInsertAfter,
        int X,
        int Y,
        int cx,
        int cy,
        uint uFlags);
}