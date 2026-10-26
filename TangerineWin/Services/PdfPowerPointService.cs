using System.IO;

using DocumentFormat.OpenXml.Packaging;

using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace TangerineWin.Services;

public sealed class PdfPowerPointService
{
    private readonly PdfRenderService
        _pdfRenderService =
            new PdfRenderService();

    public async Task<string> ConvertAsync(
        string sourcePath)
    {
        if (!File.Exists(sourcePath))
        {
            throw new FileNotFoundException(
                "PDF file was not found.",
                sourcePath);
        }

        var templatePath =
            Path.Combine(
                AppContext.BaseDirectory,
                "Assets",
                "PdfTemplate.pptx");

        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException(
                "PowerPoint template was not found.",
                templatePath);
        }

        var outputPath =
            CreateUniqueOutputPath(
                sourcePath);

        //
        // Создаём будущий PPTX как копию
        // корректного PowerPoint-шаблона.
        //
        File.Copy(
            templatePath,
            outputPath,
            overwrite: false);

        var renderedDirectory =
            await _pdfRenderService
                .RenderToImagesAsync(
                    sourcePath,
                    "pdf_to_png");

        try
        {
            var pageImages =
                Directory
                    .GetFiles(
                        renderedDirectory,
                        "*.png")
                    .OrderBy(
                        NaturalPageOrder)
                    .ToArray();

            if (pageImages.Length == 0)
            {
                throw new InvalidOperationException(
                    "PDF pages were not rendered.");
            }

            await Task.Run(() =>
            {
                FillPresentation(
                    outputPath,
                    pageImages);
            });

            return outputPath;
        }
        catch
        {
            //
            // Если создание не удалось —
            // не оставляем битый PPTX.
            //
            try
            {
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
            }
            catch
            {
            }

            throw;
        }
        finally
        {
            try
            {
                if (Directory.Exists(
                        renderedDirectory))
                {
                    Directory.Delete(
                        renderedDirectory,
                        true);
                }
            }
            catch
            {
            }
        }
    }

    // ====================================================
    // FILL EXISTING TEMPLATE
    // ====================================================

    private static void FillPresentation(
        string presentationPath,
        IReadOnlyList<string> pageImages)
    {
        using var document =
            PresentationDocument.Open(
                presentationPath,
                true);

        var presentationPart =
            document.PresentationPart
            ?? throw new InvalidOperationException(
                "Template does not contain PresentationPart.");

        var presentation =
            presentationPart.Presentation
            ?? throw new InvalidOperationException(
                "Template does not contain Presentation.");

        var slideIdList =
            presentation.SlideIdList
            ?? throw new InvalidOperationException(
                "Template does not contain SlideIdList.");

        //
        // Берём первый слайд шаблона.
        //
        var firstSlideId =
            slideIdList
                .Elements<P.SlideId>()
                .FirstOrDefault()
            ?? throw new InvalidOperationException(
                "Template must contain one blank slide.");

        var firstRelationshipId =
            firstSlideId.RelationshipId?.Value
            ?? throw new InvalidOperationException(
                "Template slide has no relationship.");

        var templateSlidePart =
            (SlidePart)
            presentationPart.GetPartById(
                firstRelationshipId);

        //
        // Нам нужен layout пустого слайда.
        //
        var slideLayoutPart =
            templateSlidePart.SlideLayoutPart
            ?? throw new InvalidOperationException(
                "Template slide does not contain SlideLayoutPart.");

        //
        // Удаляем сам шаблонный слайд,
        // но layout/master остаются.
        //
        firstSlideId.Remove();

        presentationPart.DeletePart(
            templateSlidePart);

        //
        // Определяем следующий допустимый SlideId.
        //
        uint nextSlideId =
            256;

        var existingIds =
            slideIdList
                .Elements<P.SlideId>()
                .Select(id =>
                    id.Id?.Value ?? 0)
                .ToArray();

        if (existingIds.Length > 0)
        {
            nextSlideId =
                Math.Max(
                    256,
                    existingIds.Max() + 1);
        }

        foreach (var imagePath
                 in pageImages)
        {
            var slidePart =
                presentationPart
                    .AddNewPart<SlidePart>();

            slidePart.AddPart(
                slideLayoutPart);

            CreateImageSlide(
                slidePart,
                imagePath,
                presentation);

            var relationshipId =
                presentationPart
                    .GetIdOfPart(
                        slidePart);

            slideIdList.Append(
                new P.SlideId
                {
                    Id =
                        nextSlideId++,

                    RelationshipId =
                        relationshipId
                });
        }

        presentation.Save();
    }

    // ====================================================
    // CREATE IMAGE SLIDE
    // ====================================================

    private static void CreateImageSlide(
        SlidePart slidePart,
        string imagePath,
        P.Presentation presentation)
    {
        var slideSize =
            presentation.SlideSize
            ?? throw new InvalidOperationException(
                "Template does not contain SlideSize.");

        var slideWidth =
            slideSize.Cx?.Value
            ?? throw new InvalidOperationException(
                "Template slide width is missing.");

        var slideHeight =
            slideSize.Cy?.Value
            ?? throw new InvalidOperationException(
                "Template slide height is missing.");

        var shapeTree =
            CreateEmptyShapeTree();

        var imagePart =
            slidePart.AddImagePart(
                ImagePartType.Png);

        using (var stream =
               File.OpenRead(
                   imagePath))
        {
            imagePart.FeedData(
                stream);
        }

        var imageRelationshipId =
            slidePart.GetIdOfPart(
                imagePart);

        var picture =
            new P.Picture(

                new P.NonVisualPictureProperties(

                    new P.NonVisualDrawingProperties
                    {
                        Id = 2,
                        Name =
                            Path.GetFileName(
                                imagePath)
                    },

                    new P.NonVisualPictureDrawingProperties(
                        new A.PictureLocks
                        {
                            NoChangeAspect =
                                true
                        }),

                    new P.ApplicationNonVisualDrawingProperties()),

                new P.BlipFill(

                    new A.Blip
                    {
                        Embed =
                            imageRelationshipId
                    },

                    new A.Stretch(
                        new A.FillRectangle())),

                new P.ShapeProperties(

                    new A.Transform2D(

                        new A.Offset
                        {
                            X = 0,
                            Y = 0
                        },

                        new A.Extents
                        {
                            Cx =
                                slideWidth,

                            Cy =
                                slideHeight
                        }),

                    new A.PresetGeometry(
                        new A.AdjustValueList())
                    {
                        Preset =
                            A.ShapeTypeValues.Rectangle
                    }));

        shapeTree.Append(
            picture);

        slidePart.Slide =
            new P.Slide(

                new P.CommonSlideData(
                    shapeTree),

                new P.ColorMapOverride(
                    new A.MasterColorMapping()));

        slidePart.Slide.Save();
    }

    // ====================================================
    // EMPTY SHAPE TREE
    // ====================================================

    private static P.ShapeTree
        CreateEmptyShapeTree()
    {
        return new P.ShapeTree(

            new P.NonVisualGroupShapeProperties(

                new P.NonVisualDrawingProperties
                {
                    Id = 1,
                    Name =
                        string.Empty
                },

                new P.NonVisualGroupShapeDrawingProperties(),

                new P.ApplicationNonVisualDrawingProperties()),

            new P.GroupShapeProperties(

                new A.TransformGroup(

                    new A.Offset
                    {
                        X = 0,
                        Y = 0
                    },

                    new A.Extents
                    {
                        Cx = 0,
                        Cy = 0
                    },

                    new A.ChildOffset
                    {
                        X = 0,
                        Y = 0
                    },

                    new A.ChildExtents
                    {
                        Cx = 0,
                        Cy = 0
                    })));
    }

    // ====================================================
    // PAGE ORDER
    // ====================================================

    private static int NaturalPageOrder(
        string path)
    {
        var fileName =
            Path.GetFileNameWithoutExtension(
                path);

        var digits =
            new string(
                fileName
                    .Where(
                        char.IsDigit)
                    .ToArray());

        if (int.TryParse(
                digits,
                out var page))
        {
            return page;
        }

        return int.MaxValue;
    }

    // ====================================================
    // OUTPUT
    // ====================================================

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

        var candidate =
            Path.Combine(
                directory,
                $"{name}.pptx");

        if (!File.Exists(candidate))
        {
            return candidate;
        }

        var index = 1;

        while (true)
        {
            candidate =
                Path.Combine(
                    directory,
                    $"{name} ({index}).pptx");

            if (!File.Exists(candidate))
            {
                return candidate;
            }

            index++;
        }
    }
}