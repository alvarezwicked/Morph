using System.IO;
using System.Windows;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using ImageMagick;

using WpfCanvas = System.Windows.Controls.Canvas;
using WpfMessageBox = System.Windows.MessageBox;
using WpfMouseButtonEventArgs = System.Windows.Input.MouseButtonEventArgs;
using WpfMouseEventArgs = System.Windows.Input.MouseEventArgs;
using WpfPoint = System.Windows.Point;
using WpfVisibility = System.Windows.Visibility;

namespace TangerineWin.Views;

public partial class ImageEditorDialog : Window
{
    private readonly string _sourcePath;

    private readonly DispatcherTimer _qualityPreviewTimer;

    private MagickImage? _workingImage;

    private bool _isSelecting;

    private WpfPoint _startPoint;

    private double _selectionX;
    private double _selectionY;
    private double _selectionWidth;
    private double _selectionHeight;

    private int _originalWidth;
    private int _originalHeight;

    public string? OutputPath { get; private set; }

    public ImageEditorDialog(
        string sourcePath)
    {
        InitializeComponent();

        _sourcePath = sourcePath;

        _qualityPreviewTimer =
            new DispatcherTimer
            {
                Interval =
                    TimeSpan.FromMilliseconds(180)
            };

        _qualityPreviewTimer.Tick +=
            OnQualityPreviewTimerTick;

        Loaded += OnLoaded;
        Closed += OnClosed;
    }

    private void OnLoaded(
        object sender,
        RoutedEventArgs e)
    {
        LoadOriginalImage();

        FitPreviewContainer.Visibility =
            Visibility.Visible;

        ActualSizeScrollViewer.Visibility =
            Visibility.Collapsed;

        FitButton.FontWeight =
            FontWeights.Bold;

        ActualSizeButton.FontWeight =
            FontWeights.Normal;
    }

    private void OnClosed(
        object? sender,
        EventArgs e)
    {
        _qualityPreviewTimer.Stop();

        _workingImage?.Dispose();
        _workingImage = null;
    }

    // ----------------------------------------------------
    // LOAD
    // ----------------------------------------------------

    private void LoadOriginalImage()
    {
        _workingImage?.Dispose();

        _workingImage =
            new MagickImage(
                _sourcePath);

        _workingImage.AutoOrient();

        _originalWidth =
            (int)_workingImage.Width;

        _originalHeight =
            (int)_workingImage.Height;

        ResetSelection();
        RefreshPreview();
    }

    // ----------------------------------------------------
    // PREVIEW
    // ----------------------------------------------------

    private void RefreshPreview()
    {
        if (_workingImage is null)
            return;

        RefreshMainPreview();
    }

    private void RefreshMainPreview()
    {
        if (_workingImage is null)
            return;

        var quality =
            (uint)Math.Round(
                QualitySlider.Value);

        var extension =
            Path.GetExtension(
                _sourcePath)
            .ToLowerInvariant();

        using var preview =
            (MagickImage)_workingImage.Clone();

        //
        // Здесь делаем реальное preview-кодирование
        // для JPG / WEBP.
        //
        switch (extension)
        {
            case ".jpg":
            case ".jpeg":
                {
                    preview.Format =
                        MagickFormat.Jpeg;

                    preview.Quality =
                        quality;

                    break;
                }

            case ".webp":
                {
                    preview.Format =
                        MagickFormat.WebP;

                    preview.Quality =
                        quality;

                    break;
                }

            case ".png":
                {
                    preview.Format =
                        MagickFormat.Png;

                    preview.Quality = 100;

                    break;
                }

            default:
                {
                    preview.Format =
                        MagickFormat.Png;

                    break;
                }
        }

        //
        // Важно:
        // сначала реально кодируем,
        // потом снова декодируем.
        //
        // Так в preview будут настоящие
        // JPEG / WEBP артефакты.
        //
        var encodedBytes =
            preview.ToByteArray();

        using var decoded =
            new MagickImage(
                encodedBytes);

        decoded.Format =
            MagickFormat.Png;

        var previewBytes =
            decoded.ToByteArray();

        var bitmap =
            CreateBitmapImage(
                previewBytes);

        PreviewImage.Source =
            bitmap;

        ActualSizeImage.Source =
            bitmap;

        ImageSizeText.Text =
            $"{_workingImage.Width} × {_workingImage.Height} px" +
            (
                _workingImage.Width !=
                (uint)_originalWidth ||
                _workingImage.Height !=
                (uint)_originalHeight
                    ? $"   •   исходное: {_originalWidth} × {_originalHeight}"
                    : ""
            );

        ImageOverlaySizeText.Text =
            $"{_workingImage.Width} × {_workingImage.Height} px";

        if (extension == ".png")
        {
            ImageOverlayQualityText.Text =
                "PNG • lossless";
        }
        else
        {
            ImageOverlayQualityText.Text =
                $"Quality: {quality}";
        }

        WidthTextBox.Text =
            _workingImage.Width.ToString();

        HeightTextBox.Text =
            _workingImage.Height.ToString();
    }

    private static BitmapImage CreateBitmapImage(
        byte[] bytes)
    {
        using var stream =
            new MemoryStream(
                bytes);

        var bitmap =
            new BitmapImage();

        bitmap.BeginInit();

        bitmap.CacheOption =
            BitmapCacheOption.OnLoad;

        bitmap.StreamSource =
            stream;

        bitmap.EndInit();
        bitmap.Freeze();

        return bitmap;
    }
    // ----------------------------------------------------
    // VIEW MODE
    // ----------------------------------------------------

    private void OnFitClick(
        object sender,
        RoutedEventArgs e)
    {
        FitPreviewContainer.Visibility =
            Visibility.Visible;

        ActualSizeScrollViewer.Visibility =
            Visibility.Collapsed;

        FitButton.FontWeight =
            FontWeights.Bold;

        ActualSizeButton.FontWeight =
            FontWeights.Normal;
    }

    private void OnActualSizeClick(
        object sender,
        RoutedEventArgs e)
    {
        ResetSelection();

        FitPreviewContainer.Visibility =
            Visibility.Collapsed;

        ActualSizeScrollViewer.Visibility =
            Visibility.Visible;

        FitButton.FontWeight =
            FontWeights.Normal;

        ActualSizeButton.FontWeight =
            FontWeights.Bold;
    }
    // ----------------------------------------------------
    // ROTATE
    // ----------------------------------------------------

    private void OnRotateLeftClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        _workingImage.Rotate(-90);
        _workingImage.ResetPage();

        ResetSelection();
        RefreshPreview();
    }

    private void OnRotateRightClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        _workingImage.Rotate(90);
        _workingImage.ResetPage();

        ResetSelection();
        RefreshPreview();
    }

    // ----------------------------------------------------
    // FLIP
    // ----------------------------------------------------

    private void OnFlipHorizontalClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        _workingImage.Flop();

        ResetSelection();
        RefreshPreview();
    }

    private void OnFlipVerticalClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        _workingImage.Flip();

        ResetSelection();
        RefreshPreview();
    }

    // ----------------------------------------------------
    // RESIZE
    // ----------------------------------------------------

    private void OnApplyResizeClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        if (!int.TryParse(
                WidthTextBox.Text,
                out var targetWidth) ||
            targetWidth <= 0)
        {
            ShowInfo(
                "Введите корректную ширину.");

            return;
        }

        if (!int.TryParse(
                HeightTextBox.Text,
                out var targetHeight) ||
            targetHeight <= 0)
        {
            ShowInfo(
                "Введите корректную высоту.");

            return;
        }

        var currentWidth =
            (double)_workingImage.Width;

        var currentHeight =
            (double)_workingImage.Height;

        var widthScale =
            targetWidth /
            currentWidth;

        var heightScale =
            targetHeight /
            currentHeight;

        var scale =
            Math.Min(
                widthScale,
                heightScale);

        var newWidth =
            Math.Max(
                1,
                (int)Math.Round(
                    currentWidth *
                    scale));

        var newHeight =
            Math.Max(
                1,
                (int)Math.Round(
                    currentHeight *
                    scale));

        _workingImage.Resize(
            (uint)newWidth,
            (uint)newHeight);

        ResetSelection();
        RefreshPreview();
    }

    // ----------------------------------------------------
    // CROP
    // ----------------------------------------------------

    private void OnMouseLeftButtonDown(
        object sender,
        WpfMouseButtonEventArgs e)
    {
        if (_workingImage is null)
            return;

        var bounds =
            GetDisplayedImageBounds();

        var point =
            e.GetPosition(
                SelectionCanvas);

        if (!IsPointInside(
                point,
                bounds))
        {
            return;
        }

        _isSelecting = true;

        _startPoint =
            ClampPointToImage(
                point,
                bounds);

        SelectionCanvas.CaptureMouse();

        _selectionX =
            _startPoint.X;

        _selectionY =
            _startPoint.Y;

        _selectionWidth = 0;
        _selectionHeight = 0;

        SelectionRectangle.Visibility =
            WpfVisibility.Visible;

        UpdateSelectionRectangle();
    }

    private void OnMouseMove(
        object sender,
        WpfMouseEventArgs e)
    {
        if (!_isSelecting)
            return;

        var bounds =
            GetDisplayedImageBounds();

        var point =
            ClampPointToImage(
                e.GetPosition(
                    SelectionCanvas),
                bounds);

        _selectionX =
            Math.Min(
                _startPoint.X,
                point.X);

        _selectionY =
            Math.Min(
                _startPoint.Y,
                point.Y);

        _selectionWidth =
            Math.Abs(
                point.X -
                _startPoint.X);

        _selectionHeight =
            Math.Abs(
                point.Y -
                _startPoint.Y);

        UpdateSelectionRectangle();
        UpdateSelectionSizeText();
    }

    private void OnMouseLeftButtonUp(
        object sender,
        WpfMouseButtonEventArgs e)
    {
        if (!_isSelecting)
            return;

        _isSelecting = false;

        SelectionCanvas.ReleaseMouseCapture();

        if (_selectionWidth < 3 ||
            _selectionHeight < 3)
        {
            ResetSelection();

            return;
        }

        UpdateSelectionSizeText();
    }

    private void OnApplyCropClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        if (!TryGetCropPixels(
                out var x,
                out var y,
                out var width,
                out var height))
        {
            ShowInfo(
                "Сначала выделите область изображения.");

            return;
        }

        _workingImage.Crop(
            new MagickGeometry(
                x,
                y,
                (uint)width,
                (uint)height));

        _workingImage.ResetPage();

        ResetSelection();
        RefreshPreview();
    }

    // ----------------------------------------------------
    // QUALITY
    // ----------------------------------------------------

    private void OnQualityChanged(
        object sender,
        RoutedPropertyChangedEventArgs<double> e)
    {
        if (QualityText is null)
            return;

        var quality =
            (int)Math.Round(
                e.NewValue);

        QualityText.Text =
            quality.ToString();

        if (ImageOverlayQualityText is not null)
        {
            var extension =
                Path.GetExtension(
                    _sourcePath)
                .ToLowerInvariant();

            ImageOverlayQualityText.Text =
                extension == ".png"
                    ? "PNG • lossless"
                    : $"Quality: {quality}";
        }

        if (_qualityPreviewTimer is null)
            return;

        //
        // Пока ползунок двигается,
        // постоянно откладываем тяжёлый preview.
        //
        _qualityPreviewTimer.Stop();
        _qualityPreviewTimer.Start();
    }

    private void OnQualityPreviewTimerTick(
        object? sender,
        EventArgs e)
    {
        _qualityPreviewTimer.Stop();

        RefreshMainPreview();
    }

    // ----------------------------------------------------
    // RESET
    // ----------------------------------------------------

    private void OnResetClick(
        object sender,
        RoutedEventArgs e)
    {
        _workingImage?.Dispose();

        _workingImage =
            new MagickImage(
                _sourcePath);

        _workingImage.AutoOrient();

        ResetSelection();

        QualitySlider.Value = 85;

        RemoveMetadataCheckBox.IsChecked =
            false;

        RefreshPreview();
    }

    // ----------------------------------------------------
    // SAVE
    // ----------------------------------------------------

    private async void OnSaveClick(
        object sender,
        RoutedEventArgs e)
    {
        if (_workingImage is null)
            return;

        var quality =
            (uint)Math.Round(
                QualitySlider.Value);

        var removeMetadata =
            RemoveMetadataCheckBox.IsChecked ==
            true;

        var extension =
            Path.GetExtension(
                _sourcePath)
            .ToLowerInvariant();

        var outputPath =
            CreateUniqueOutputPath(
                _sourcePath);

        var image =
            (MagickImage)_workingImage.Clone();

        try
        {
            await Task.Run(() =>
            {
                if (removeMetadata)
                {
                    image.Strip();
                }

                switch (extension)
                {
                    case ".jpg":
                    case ".jpeg":
                        {
                            image.Format =
                                MagickFormat.Jpeg;

                            image.Quality =
                                quality;

                            break;
                        }

                    case ".webp":
                        {
                            image.Format =
                                MagickFormat.WebP;

                            image.Quality =
                                quality;

                            break;
                        }

                    case ".png":
                        {
                            image.Format =
                                MagickFormat.Png;

                            image.Quality = 100;

                            break;
                        }

                    default:
                        {
                            throw new NotSupportedException(
                                $"Unsupported image format: {extension}");
                        }
                }

                image.Write(
                    outputPath);
            });

            OutputPath =
                outputPath;

            DialogResult = true;
        }
        catch (Exception ex)
        {
            WpfMessageBox.Show(
                ex.Message,
                "Morph",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            image.Dispose();
        }
    }

    // ----------------------------------------------------
    // CROP COORDINATES
    // ----------------------------------------------------

    private bool TryGetCropPixels(
        out int x,
        out int y,
        out int width,
        out int height)
    {
        x = 0;
        y = 0;
        width = 0;
        height = 0;

        if (_workingImage is null)
            return false;

        if (_selectionWidth <= 0 ||
            _selectionHeight <= 0)
        {
            return false;
        }

        var bounds =
            GetDisplayedImageBounds();

        if (bounds.Width <= 0 ||
            bounds.Height <= 0)
        {
            return false;
        }

        var relativeX =
            _selectionX -
            bounds.X;

        var relativeY =
            _selectionY -
            bounds.Y;

        var scaleX =
            (double)_workingImage.Width /
            bounds.Width;

        var scaleY =
            (double)_workingImage.Height /
            bounds.Height;

        x =
            (int)Math.Round(
                relativeX *
                scaleX);

        y =
            (int)Math.Round(
                relativeY *
                scaleY);

        width =
            (int)Math.Round(
                _selectionWidth *
                scaleX);

        height =
            (int)Math.Round(
                _selectionHeight *
                scaleY);

        x =
            Math.Clamp(
                x,
                0,
                Math.Max(
                    0,
                    (int)_workingImage.Width - 1));

        y =
            Math.Clamp(
                y,
                0,
                Math.Max(
                    0,
                    (int)_workingImage.Height - 1));

        width =
            Math.Clamp(
                width,
                1,
                (int)_workingImage.Width - x);

        height =
            Math.Clamp(
                height,
                1,
                (int)_workingImage.Height - y);

        return true;
    }

    private void UpdateSelectionSizeText()
    {
        if (TryGetCropPixels(
                out _,
                out _,
                out var width,
                out var height))
        {
            SelectionSizeText.Text =
                $"{width} × {height} px";
        }
        else
        {
            SelectionSizeText.Text =
                "Область не выбрана";
        }
    }

    private void UpdateSelectionRectangle()
    {
        WpfCanvas.SetLeft(
            SelectionRectangle,
            _selectionX);

        WpfCanvas.SetTop(
            SelectionRectangle,
            _selectionY);

        SelectionRectangle.Width =
            _selectionWidth;

        SelectionRectangle.Height =
            _selectionHeight;
    }

    private void ResetSelection()
    {
        _isSelecting = false;

        SelectionCanvas.ReleaseMouseCapture();

        _selectionX = 0;
        _selectionY = 0;
        _selectionWidth = 0;
        _selectionHeight = 0;

        SelectionRectangle.Visibility =
            WpfVisibility.Collapsed;

        SelectionSizeText.Text =
            "Область не выбрана";
    }

    // ----------------------------------------------------
    // DISPLAYED IMAGE BOUNDS
    // ----------------------------------------------------

    private ImageBounds GetDisplayedImageBounds()
    {
        if (_workingImage is null)
        {
            return new ImageBounds(
                0,
                0,
                0,
                0);
        }

        var availableWidth =
            SelectionCanvas.ActualWidth;

        var availableHeight =
            SelectionCanvas.ActualHeight;

        if (availableWidth <= 0 ||
            availableHeight <= 0)
        {
            return new ImageBounds(
                0,
                0,
                0,
                0);
        }

        var imageWidth =
            (double)_workingImage.Width;

        var imageHeight =
            (double)_workingImage.Height;

        var scale =
            Math.Min(
                availableWidth /
                imageWidth,

                availableHeight /
                imageHeight);

        var displayedWidth =
            imageWidth *
            scale;

        var displayedHeight =
            imageHeight *
            scale;

        var x =
            (
                availableWidth -
                displayedWidth
            ) / 2.0;

        var y =
            (
                availableHeight -
                displayedHeight
            ) / 2.0;

        return new ImageBounds(
            x,
            y,
            displayedWidth,
            displayedHeight);
    }

    private static bool IsPointInside(
        WpfPoint point,
        ImageBounds bounds)
    {
        return
            point.X >= bounds.X &&
            point.X <= bounds.X + bounds.Width &&
            point.Y >= bounds.Y &&
            point.Y <= bounds.Y + bounds.Height;
    }

    private static WpfPoint ClampPointToImage(
        WpfPoint point,
        ImageBounds bounds)
    {
        return new WpfPoint(
            Math.Clamp(
                point.X,
                bounds.X,
                bounds.X + bounds.Width),

            Math.Clamp(
                point.Y,
                bounds.Y,
                bounds.Y + bounds.Height));
    }

    // ----------------------------------------------------
    // OUTPUT
    // ----------------------------------------------------

    private static string CreateUniqueOutputPath(
        string sourcePath)
    {
        var directory =
            Path.GetDirectoryName(
                sourcePath)
            ?? throw new InvalidOperationException(
                "Unable to determine source directory.");

        var name =
            Path.GetFileNameWithoutExtension(
                sourcePath);

        var extension =
            Path.GetExtension(
                sourcePath);

        var candidate =
            Path.Combine(
                directory,
                $"{name}-edited{extension}");

        if (!File.Exists(candidate))
            return candidate;

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name}-edited ({index}){extension}");

            if (!File.Exists(candidate))
                return candidate;

            index++;
        }
    }

    private static void ShowInfo(
        string message)
    {
        WpfMessageBox.Show(
            message,
            "Morph",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private sealed record ImageBounds(
        double X,
        double Y,
        double Width,
        double Height);
}