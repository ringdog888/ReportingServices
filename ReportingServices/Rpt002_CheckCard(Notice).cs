using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using V = DocumentFormat.OpenXml.Vml;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using M = DocumentFormat.OpenXml.Math;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using System.Data;

namespace Rpt002_Notice
{
    public class GeneratedClass
    {
        //Creates Report Tool
        ReportingServices.RptTool RptTool = new ReportingServices.RptTool();

        // Data Source
        public DataTable dt { get; set; }
        public string[] ChineseW = {"○", "一", "二", "三", "四", "五", "六", "七", "八", "九", "十"};

        // Creates a WordprocessingDocument.
        public void CreatePackage(string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId8");
            GenerateFontTablePart1Content(fontTablePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId3");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            ImagePart imagePart1 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rId7");
            GenerateImagePart1Content(imagePart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId2");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            ImagePart imagePart2 = mainDocumentPart1.AddNewPart<ImagePart>("image/png", "rId10");
            GenerateImagePart2Content(imagePart2);

            CustomXmlPart customXmlPart1 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId1");
            GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId6");
            GenerateEndnotesPart1Content(endnotesPart1);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId5");
            GenerateFootnotesPart1Content(footnotesPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId4");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId9");
            GenerateThemePart1Content(themePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal.dotm";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "7";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "2";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "57";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "327";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "2";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "Crystal Decisions";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "383";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of mainDocumentPart1.
        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            Document document1 = new Document();
            document1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Body body1 = new Body();

            Paragraph paragraph50 = new Paragraph() { RsidParagraphMarkRevision = "006369CA", RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "006369CA", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            WidowControl widowControl2 = new WidowControl();

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            RunFonts runFonts186 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern186 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript186 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties50.Append(runFonts186);
            paragraphMarkRunProperties50.Append(kern186);
            paragraphMarkRunProperties50.Append(fontSizeComplexScript186);

            paragraphProperties50.Append(widowControl2);
            paragraphProperties50.Append(paragraphMarkRunProperties50);

            paragraph50.Append(paragraphProperties50);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "006369CA", RsidR = "00BB0351", RsidSect = "000F2B42" };
            SectionType sectionType1 = new SectionType() { Val = SectionMarkValues.Continuous };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)16560U };
            PageMargin pageMargin1 = new PageMargin() { Top = 360, Right = (UInt32Value)360U, Bottom = 360, Left = (UInt32Value)360U, Header = (UInt32Value)720U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "720" };
            NoEndnote noEndnote1 = new NoEndnote();

            sectionProperties1.Append(sectionType1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(noEndnote1);

            //-------
            for (int i = 0; i < dt.Rows.Count; i++) {
                body1 = AddConten(body1,i);

                //除最後一筆資料外，印分頁符號
                if (i < dt.Rows.Count - 1)
                {
                    body1.Append(RptTool.GetBreakValues());
                }
            }






            body1.Append(paragraph50);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
        }


        Body AddConten(Body body1,int i)
        {
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "000F2B42" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            FrameProperties frameProperties1 = new FrameProperties() { Width = "3120", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "7321", Y = "1561", HeightType = HeightRuleValues.Exact };

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            tabs1.Append(tabStop3);
            tabs1.Append(tabStop4);
            tabs1.Append(tabStop5);
            tabs1.Append(tabStop6);
            tabs1.Append(tabStop7);
            tabs1.Append(tabStop8);
            AutoSpaceDE autoSpaceDE1 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN1 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent1 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern1 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(kern1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(frameProperties1);
            paragraphProperties1.Append(tabs1);
            paragraphProperties1.Append(autoSpaceDE1);
            paragraphProperties1.Append(autoSpaceDN1);
            paragraphProperties1.Append(adjustRightIndent1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "000F2B42" };

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Picture picture1 = new Picture();

            V.Line line1 = new V.Line() { Id = "_x0000_s1045", Style = "position:absolute;z-index:-251637760;mso-position-horizontal-relative:page;mso-position-vertical-relative:page", AllowInCell = false, StrokeWeight = "1pt", From = "60pt,30pt", To = "60pt,708.85pt" };
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

            line1.Append(textWrap1);

            picture1.Append(line1);

            run1.Append(runProperties1);
            run1.Append(picture1);

            Run run2 = new Run() { RsidRunProperties = "000F2B42" };

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);

            Picture picture2 = new Picture();

            V.Line line2 = new V.Line() { Id = "_x0000_s1046", Style = "position:absolute;z-index:-251636736;mso-position-horizontal-relative:page;mso-position-vertical-relative:page", AllowInCell = false, StrokeWeight = "1pt", From = "60pt,30pt", To = "540.05pt,30pt" };
            Wvml.TextWrap textWrap2 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

            line2.Append(textWrap2);

            picture2.Append(line2);

            run2.Append(runProperties2);
            run2.Append(picture2);

            Run run3 = new Run() { RsidRunProperties = "000F2B42" };

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);

            Picture picture3 = new Picture();

            V.Line line3 = new V.Line() { Id = "_x0000_s1047", Style = "position:absolute;z-index:-251635712;mso-position-horizontal-relative:page;mso-position-vertical-relative:page", AllowInCell = false, StrokeWeight = "1pt", From = "540pt,30pt", To = "540pt,708.85pt" };
            Wvml.TextWrap textWrap3 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

            line3.Append(textWrap3);

            picture3.Append(line3);

            run3.Append(runProperties3);
            run3.Append(picture3);

            Run run4 = new Run() { RsidRunProperties = "000F2B42" };

            RunProperties runProperties4 = new RunProperties();
            NoProof noProof4 = new NoProof();

            runProperties4.Append(noProof4);

            Picture picture4 = new Picture();

            V.Line line4 = new V.Line() { Id = "_x0000_s1048", Style = "position:absolute;z-index:-251634688;mso-position-horizontal-relative:page;mso-position-vertical-relative:page", AllowInCell = false, StrokeWeight = "1pt", From = "60pt,107.25pt", To = "540.05pt,107.25pt" };
            Wvml.TextWrap textWrap4 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

            line4.Append(textWrap4);

            picture4.Append(line4);

            run4.Append(runProperties4);
            run4.Append(picture4);

            Run run5 = new Run() { RsidRunProperties = "000F2B42" };

            RunProperties runProperties5 = new RunProperties();
            NoProof noProof5 = new NoProof();

            runProperties5.Append(noProof5);

            Picture picture5 = new Picture();

            V.Line line5 = new V.Line() { Id = "_x0000_s1049", Style = "position:absolute;z-index:-251633664;mso-position-horizontal-relative:page;mso-position-vertical-relative:page", AllowInCell = false, StrokeWeight = "1pt", From = "60pt,708.8pt", To = "540.05pt,708.8pt" };
            Wvml.TextWrap textWrap5 = new Wvml.TextWrap() { AnchorX = Wvml.HorizontalAnchorValues.Page, AnchorY = Wvml.VerticalAnchorValues.Page };

            line5.Append(textWrap5);

            picture5.Append(line5);

            run5.Append(runProperties5);
            run5.Append(picture5);

            Run run6 = new Run() { RsidRunAddition = "00BB0351" };

            RunProperties runProperties6 = new RunProperties();
            NoProof noProof6 = new NoProof();

            runProperties6.Append(noProof6);

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251683840U, BehindDoc = true, Locked = false, LayoutInCell = false, AllowOverlap = true };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Page };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "1066800";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Page };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "533400";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 2759075L, Cy = 674370L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 3175L, BottomEdge = 0L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)26U, Name = "圖片 26" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture6 = new Pic.Picture();
            picture6.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 26" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();
            A.Blip blip1 = new A.Blip() { Embed = "rId7" };
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle();

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 2759075L, Cy = 674370L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);

            picture6.Append(nonVisualPictureProperties1);
            picture6.Append(blipFill1);
            picture6.Append(shapeProperties1);

            graphicData1.Append(picture6);

            graphic1.Append(graphicData1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);

            drawing1.Append(anchor1);

            run6.Append(runProperties6);
            run6.Append(drawing1);

            Run run7 = new Run() { RsidRunAddition = "00BB0351" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Bold bold1 = new Bold();
            Color color1 = new Color() { Val = "000000" };
            Kern kern2 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1 = new FontSize() { Val = "35" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

            runProperties7.Append(runFonts2);
            runProperties7.Append(bold1);
            runProperties7.Append(color1);
            runProperties7.Append(kern2);
            runProperties7.Append(fontSize1);
            runProperties7.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = "查核證 ( 通知 )";

            run7.Append(runProperties7);
            run7.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);
            paragraph1.Append(run4);
            paragraph1.Append(run5);
            paragraph1.Append(run6);
            paragraph1.Append(run7);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            FrameProperties frameProperties2 = new FrameProperties() { Width = "1920", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1681", Y = "2777", HeightType = HeightRuleValues.Exact };

            Tabs tabs2 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };

            tabs2.Append(tabStop9);
            tabs2.Append(tabStop10);
            tabs2.Append(tabStop11);
            tabs2.Append(tabStop12);
            tabs2.Append(tabStop13);
            AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern3 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(kern3);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

            paragraphProperties2.Append(frameProperties2);
            paragraphProperties2.Append(tabs2);
            paragraphProperties2.Append(autoSpaceDE2);
            paragraphProperties2.Append(autoSpaceDN2);
            paragraphProperties2.Append(adjustRightIndent2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color2 = new Color() { Val = "000000" };
            Kern kern4 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize2 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(runFonts4);
            runProperties8.Append(color2);
            runProperties8.Append(kern4);
            runProperties8.Append(fontSize2);
            runProperties8.Append(fontSizeComplexScript4);
            Text text2 = new Text();
            text2.Text = "受";

            run8.Append(runProperties8);
            run8.Append(text2);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color3 = new Color() { Val = "000000" };
            Kern kern5 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize3 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            runProperties9.Append(runFonts5);
            runProperties9.Append(color3);
            runProperties9.Append(kern5);
            runProperties9.Append(fontSize3);
            runProperties9.Append(fontSizeComplexScript5);
            Text text3 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text3.Text = "　";

            run9.Append(runProperties9);
            run9.Append(text3);

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color4 = new Color() { Val = "000000" };
            Kern kern6 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize4 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            runProperties10.Append(runFonts6);
            runProperties10.Append(color4);
            runProperties10.Append(kern6);
            runProperties10.Append(fontSize4);
            runProperties10.Append(fontSizeComplexScript6);
            Text text4 = new Text();
            text4.Text = "文";

            run10.Append(runProperties10);
            run10.Append(text4);

            Run run11 = new Run();

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color5 = new Color() { Val = "000000" };
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize5 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            runProperties11.Append(runFonts7);
            runProperties11.Append(color5);
            runProperties11.Append(kern7);
            runProperties11.Append(fontSize5);
            runProperties11.Append(fontSizeComplexScript7);
            Text text5 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text5.Text = "　";

            run11.Append(runProperties11);
            run11.Append(text5);

            Run run12 = new Run();

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color6 = new Color() { Val = "000000" };
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize6 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

            runProperties12.Append(runFonts8);
            runProperties12.Append(color6);
            runProperties12.Append(kern8);
            runProperties12.Append(fontSize6);
            runProperties12.Append(fontSizeComplexScript8);
            Text text6 = new Text();
            text6.Text = "者";

            run12.Append(runProperties12);
            run12.Append(text6);

            Run run13 = new Run();

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color7 = new Color() { Val = "000000" };
            Kern kern9 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

            runProperties13.Append(runFonts9);
            runProperties13.Append(color7);
            runProperties13.Append(kern9);
            runProperties13.Append(fontSize7);
            runProperties13.Append(fontSizeComplexScript9);
            Text text7 = new Text();
            text7.Text = "：";

            run13.Append(runProperties13);
            run13.Append(text7);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run8);
            paragraph2.Append(run9);
            paragraph2.Append(run10);
            paragraph2.Append(run11);
            paragraph2.Append(run12);
            paragraph2.Append(run13);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            FrameProperties frameProperties3 = new FrameProperties() { Width = "2520", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1681", Y = "3497", HeightType = HeightRuleValues.Exact };

            Tabs tabs3 = new Tabs();
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };

            tabs3.Append(tabStop14);
            tabs3.Append(tabStop15);
            tabs3.Append(tabStop16);
            tabs3.Append(tabStop17);
            tabs3.Append(tabStop18);
            tabs3.Append(tabStop19);
            AutoSpaceDE autoSpaceDE3 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN3 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent3 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern10 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties3.Append(runFonts10);
            paragraphMarkRunProperties3.Append(kern10);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript10);

            paragraphProperties3.Append(frameProperties3);
            paragraphProperties3.Append(tabs3);
            paragraphProperties3.Append(autoSpaceDE3);
            paragraphProperties3.Append(autoSpaceDN3);
            paragraphProperties3.Append(adjustRightIndent3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run14 = new Run();

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color8 = new Color() { Val = "000000" };
            Kern kern11 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize8 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

            runProperties14.Append(runFonts11);
            runProperties14.Append(color8);
            runProperties14.Append(kern11);
            runProperties14.Append(fontSize8);
            runProperties14.Append(fontSizeComplexScript11);
            Text text8 = new Text();
            text8.Text = "查";

            run14.Append(runProperties14);
            run14.Append(text8);

            Run run15 = new Run();

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color9 = new Color() { Val = "000000" };
            Kern kern12 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize9 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            runProperties15.Append(runFonts12);
            runProperties15.Append(color9);
            runProperties15.Append(kern12);
            runProperties15.Append(fontSize9);
            runProperties15.Append(fontSizeComplexScript12);
            Text text9 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text9.Text = "　";

            run15.Append(runProperties15);
            run15.Append(text9);

            Run run16 = new Run();

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color10 = new Color() { Val = "000000" };
            Kern kern13 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize10 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

            runProperties16.Append(runFonts13);
            runProperties16.Append(color10);
            runProperties16.Append(kern13);
            runProperties16.Append(fontSize10);
            runProperties16.Append(fontSizeComplexScript13);
            Text text10 = new Text();
            text10.Text = "核";

            run16.Append(runProperties16);
            run16.Append(text10);

            Run run17 = new Run();

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color11 = new Color() { Val = "000000" };
            Kern kern14 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize11 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

            runProperties17.Append(runFonts14);
            runProperties17.Append(color11);
            runProperties17.Append(kern14);
            runProperties17.Append(fontSize11);
            runProperties17.Append(fontSizeComplexScript14);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = "　";

            run17.Append(runProperties17);
            run17.Append(text11);

            Run run18 = new Run();

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color12 = new Color() { Val = "000000" };
            Kern kern15 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize12 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

            runProperties18.Append(runFonts15);
            runProperties18.Append(color12);
            runProperties18.Append(kern15);
            runProperties18.Append(fontSize12);
            runProperties18.Append(fontSizeComplexScript15);
            Text text12 = new Text();
            text12.Text = "事";

            run18.Append(runProperties18);
            run18.Append(text12);

            Run run19 = new Run();

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color13 = new Color() { Val = "000000" };
            Kern kern16 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize13 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

            runProperties19.Append(runFonts16);
            runProperties19.Append(color13);
            runProperties19.Append(kern16);
            runProperties19.Append(fontSize13);
            runProperties19.Append(fontSizeComplexScript16);
            Text text13 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text13.Text = "　";

            run19.Append(runProperties19);
            run19.Append(text13);

            Run run20 = new Run();

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color14 = new Color() { Val = "000000" };
            Kern kern17 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize14 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

            runProperties20.Append(runFonts17);
            runProperties20.Append(color14);
            runProperties20.Append(kern17);
            runProperties20.Append(fontSize14);
            runProperties20.Append(fontSizeComplexScript17);
            Text text14 = new Text();
            text14.Text = "由";

            run20.Append(runProperties20);
            run20.Append(text14);

            Run run21 = new Run();

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color15 = new Color() { Val = "000000" };
            Kern kern18 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize15 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

            runProperties21.Append(runFonts18);
            runProperties21.Append(color15);
            runProperties21.Append(kern18);
            runProperties21.Append(fontSize15);
            runProperties21.Append(fontSizeComplexScript18);
            Text text15 = new Text();
            text15.Text = "：";

            run21.Append(runProperties21);
            run21.Append(text15);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run14);
            paragraph3.Append(run15);
            paragraph3.Append(run16);
            paragraph3.Append(run17);
            paragraph3.Append(run18);
            paragraph3.Append(run19);
            paragraph3.Append(run20);
            paragraph3.Append(run21);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            FrameProperties frameProperties4 = new FrameProperties() { Width = "840", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "8041", Y = "2777", HeightType = HeightRuleValues.Exact };

            Tabs tabs4 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Left, Position = 720 };

            tabs4.Append(tabStop20);
            tabs4.Append(tabStop21);
            AutoSpaceDE autoSpaceDE4 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN4 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent4 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern19 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties4.Append(runFonts19);
            paragraphMarkRunProperties4.Append(kern19);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript19);

            paragraphProperties4.Append(frameProperties4);
            paragraphProperties4.Append(tabs4);
            paragraphProperties4.Append(autoSpaceDE4);
            paragraphProperties4.Append(autoSpaceDN4);
            paragraphProperties4.Append(adjustRightIndent4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run22 = new Run();

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color16 = new Color() { Val = "000000" };
            Kern kern20 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize16 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

            runProperties22.Append(runFonts20);
            runProperties22.Append(color16);
            runProperties22.Append(kern20);
            runProperties22.Append(fontSize16);
            runProperties22.Append(fontSizeComplexScript20);
            Text text16 = new Text();
            text16.Text = "）No.";

            run22.Append(runProperties22);
            run22.Append(text16);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run22);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            FrameProperties frameProperties5 = new FrameProperties() { Width = "810", Height = (UInt32Value)330U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "2881", Y = "4217", HeightType = HeightRuleValues.Exact };

            Tabs tabs5 = new Tabs();
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Left, Position = 720 };

            tabs5.Append(tabStop22);
            tabs5.Append(tabStop23);
            AutoSpaceDE autoSpaceDE5 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN5 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent5 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts21 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern21 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties5.Append(runFonts21);
            paragraphMarkRunProperties5.Append(kern21);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript21);

            paragraphProperties5.Append(frameProperties5);
            paragraphProperties5.Append(tabs5);
            paragraphProperties5.Append(autoSpaceDE5);
            paragraphProperties5.Append(autoSpaceDN5);
            paragraphProperties5.Append(adjustRightIndent5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run23 = new Run();

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color17 = new Color() { Val = "000000" };
            Kern kern22 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize17 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

            runProperties23.Append(runFonts22);
            runProperties23.Append(color17);
            runProperties23.Append(kern22);
            runProperties23.Append(fontSize17);
            runProperties23.Append(fontSizeComplexScript22);
            Text text17 = new Text();
            text17.Text = "茲";

            run23.Append(runProperties23);
            run23.Append(text17);

            Run run24 = new Run();

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color18 = new Color() { Val = "000000" };
            Kern kern23 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize18 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

            runProperties24.Append(runFonts23);
            runProperties24.Append(color18);
            runProperties24.Append(kern23);
            runProperties24.Append(fontSize18);
            runProperties24.Append(fontSizeComplexScript23);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = " ";

            run24.Append(runProperties24);
            run24.Append(text18);

            Run run25 = new Run();

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts24 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color19 = new Color() { Val = "000000" };
            Kern kern24 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize19 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };

            runProperties25.Append(runFonts24);
            runProperties25.Append(color19);
            runProperties25.Append(kern24);
            runProperties25.Append(fontSize19);
            runProperties25.Append(fontSizeComplexScript24);
            Text text19 = new Text();
            text19.Text = "派";

            run25.Append(runProperties25);
            run25.Append(text19);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run23);
            paragraph5.Append(run24);
            paragraph5.Append(run25);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            FrameProperties frameProperties6 = new FrameProperties() { Width = "720", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "3001", Y = "8537", HeightType = HeightRuleValues.Exact };

            Tabs tabs6 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Left, Position = 360 };

            tabs6.Append(tabStop24);
            AutoSpaceDE autoSpaceDE6 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN6 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent6 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern25 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties6.Append(runFonts25);
            paragraphMarkRunProperties6.Append(kern25);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript25);

            paragraphProperties6.Append(frameProperties6);
            paragraphProperties6.Append(tabs6);
            paragraphProperties6.Append(autoSpaceDE6);
            paragraphProperties6.Append(autoSpaceDN6);
            paragraphProperties6.Append(adjustRightIndent6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run26 = new Run();

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color20 = new Color() { Val = "000000" };
            Kern kern26 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize20 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

            runProperties26.Append(runFonts26);
            runProperties26.Append(color20);
            runProperties26.Append(kern26);
            runProperties26.Append(fontSize20);
            runProperties26.Append(fontSizeComplexScript26);
            Text text20 = new Text();
            text20.Text = "前赴";

            run26.Append(runProperties26);
            run26.Append(text20);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run26);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            FrameProperties frameProperties7 = new FrameProperties() { Width = "6480", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1801", Y = "9377", HeightType = HeightRuleValues.Exact };

            Tabs tabs7 = new Tabs();
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop26 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop27 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop28 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop29 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop30 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop31 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop32 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop33 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop34 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop35 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop36 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop37 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };
            TabStop tabStop38 = new TabStop() { Val = TabStopValues.Left, Position = 5040 };
            TabStop tabStop39 = new TabStop() { Val = TabStopValues.Left, Position = 5400 };
            TabStop tabStop40 = new TabStop() { Val = TabStopValues.Left, Position = 5760 };
            TabStop tabStop41 = new TabStop() { Val = TabStopValues.Left, Position = 6120 };

            tabs7.Append(tabStop25);
            tabs7.Append(tabStop26);
            tabs7.Append(tabStop27);
            tabs7.Append(tabStop28);
            tabs7.Append(tabStop29);
            tabs7.Append(tabStop30);
            tabs7.Append(tabStop31);
            tabs7.Append(tabStop32);
            tabs7.Append(tabStop33);
            tabs7.Append(tabStop34);
            tabs7.Append(tabStop35);
            tabs7.Append(tabStop36);
            tabs7.Append(tabStop37);
            tabs7.Append(tabStop38);
            tabs7.Append(tabStop39);
            tabs7.Append(tabStop40);
            tabs7.Append(tabStop41);
            AutoSpaceDE autoSpaceDE7 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN7 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent7 = new AdjustRightIndent() { Val = false };
            Indentation indentation1 = new Indentation() { FirstLine = "800" };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern27 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties7.Append(runFonts27);
            paragraphMarkRunProperties7.Append(kern27);
            paragraphMarkRunProperties7.Append(fontSizeComplexScript27);

            paragraphProperties7.Append(frameProperties7);
            paragraphProperties7.Append(tabs7);
            paragraphProperties7.Append(autoSpaceDE7);
            paragraphProperties7.Append(autoSpaceDN7);
            paragraphProperties7.Append(adjustRightIndent7);
            paragraphProperties7.Append(indentation1);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run27 = new Run();

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color21 = new Color() { Val = "000000" };
            Kern kern28 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize21 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            runProperties27.Append(runFonts28);
            runProperties27.Append(color21);
            runProperties27.Append(kern28);
            runProperties27.Append(fontSize21);
            runProperties27.Append(fontSizeComplexScript28);
            Text text21 = new Text();
            text21.Text = "貴單位查核";

            run27.Append(runProperties27);
            run27.Append(text21);

            Run run28 = new Run();

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color22 = new Color() { Val = "000000" };
            Kern kern29 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize22 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "24" };

            runProperties28.Append(runFonts29);
            runProperties28.Append(color22);
            runProperties28.Append(kern29);
            runProperties28.Append(fontSize22);
            runProperties28.Append(fontSizeComplexScript29);
            Text text22 = new Text();
            text22.Text = "，";

            run28.Append(runProperties28);
            run28.Append(text22);

            Run run29 = new Run();

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color23 = new Color() { Val = "000000" };
            Kern kern30 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize23 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            runProperties29.Append(runFonts30);
            runProperties29.Append(color23);
            runProperties29.Append(kern30);
            runProperties29.Append(fontSize23);
            runProperties29.Append(fontSizeComplexScript30);
            Text text23 = new Text();
            text23.Text = "請";

            run29.Append(runProperties29);
            run29.Append(text23);

            Run run30 = new Run();

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color24 = new Color() { Val = "000000" };
            Kern kern31 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize24 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

            runProperties30.Append(runFonts31);
            runProperties30.Append(color24);
            runProperties30.Append(kern31);
            runProperties30.Append(fontSize24);
            runProperties30.Append(fontSizeComplexScript31);
            Text text24 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text24.Text = "　";

            run30.Append(runProperties30);
            run30.Append(text24);

            Run run31 = new Run();

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color25 = new Color() { Val = "000000" };
            Kern kern32 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize25 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "24" };

            runProperties31.Append(runFonts32);
            runProperties31.Append(color25);
            runProperties31.Append(kern32);
            runProperties31.Append(fontSize25);
            runProperties31.Append(fontSizeComplexScript32);
            Text text25 = new Text();
            text25.Text = "查照並予協助為荷";

            run31.Append(runProperties31);
            run31.Append(text25);

            Run run32 = new Run();

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color26 = new Color() { Val = "000000" };
            Kern kern33 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize26 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "24" };

            runProperties32.Append(runFonts33);
            runProperties32.Append(color26);
            runProperties32.Append(kern33);
            runProperties32.Append(fontSize26);
            runProperties32.Append(fontSizeComplexScript33);
            Text text26 = new Text();
            text26.Text = "。";

            run32.Append(runProperties32);
            run32.Append(text26);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run27);
            paragraph7.Append(run28);
            paragraph7.Append(run29);
            paragraph7.Append(run30);
            paragraph7.Append(run31);
            paragraph7.Append(run32);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            FrameProperties frameProperties8 = new FrameProperties() { Width = "3000", Height = (UInt32Value)480U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "4441", Y = "11297", HeightType = HeightRuleValues.Exact };

            Tabs tabs8 = new Tabs();
            TabStop tabStop42 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop43 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop44 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop45 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop46 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop47 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop48 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop49 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };

            tabs8.Append(tabStop42);
            tabs8.Append(tabStop43);
            tabs8.Append(tabStop44);
            tabs8.Append(tabStop45);
            tabs8.Append(tabStop46);
            tabs8.Append(tabStop47);
            tabs8.Append(tabStop48);
            tabs8.Append(tabStop49);
            AutoSpaceDE autoSpaceDE8 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN8 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent8 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern34 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties8.Append(runFonts34);
            paragraphMarkRunProperties8.Append(kern34);
            paragraphMarkRunProperties8.Append(fontSizeComplexScript34);

            paragraphProperties8.Append(frameProperties8);
            paragraphProperties8.Append(tabs8);
            paragraphProperties8.Append(autoSpaceDE8);
            paragraphProperties8.Append(autoSpaceDN8);
            paragraphProperties8.Append(adjustRightIndent8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run33 = new Run();

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Bold bold2 = new Bold();
            Color color27 = new Color() { Val = "000000" };
            Kern kern35 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize27 = new FontSize() { Val = "47" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "24" };

            runProperties33.Append(runFonts35);
            runProperties33.Append(bold2);
            runProperties33.Append(color27);
            runProperties33.Append(kern35);
            runProperties33.Append(fontSize27);
            runProperties33.Append(fontSizeComplexScript35);
            Text text27 = new Text();
            text27.Text = "董事會稽核處";

            run33.Append(runProperties33);
            run33.Append(text27);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run33);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            FrameProperties frameProperties9 = new FrameProperties() { Width = "2280", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1441", Y = "13697", HeightType = HeightRuleValues.Exact };

            Tabs tabs9 = new Tabs();
            TabStop tabStop50 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop51 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop52 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop53 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop54 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop55 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };

            tabs9.Append(tabStop50);
            tabs9.Append(tabStop51);
            tabs9.Append(tabStop52);
            tabs9.Append(tabStop53);
            tabs9.Append(tabStop54);
            tabs9.Append(tabStop55);
            AutoSpaceDE autoSpaceDE9 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN9 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent9 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern36 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties9.Append(runFonts36);
            paragraphMarkRunProperties9.Append(kern36);
            paragraphMarkRunProperties9.Append(fontSizeComplexScript36);

            paragraphProperties9.Append(frameProperties9);
            paragraphProperties9.Append(tabs9);
            paragraphProperties9.Append(autoSpaceDE9);
            paragraphProperties9.Append(autoSpaceDN9);
            paragraphProperties9.Append(adjustRightIndent9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run34 = new Run();

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color28 = new Color() { Val = "000000" };
            Kern kern37 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize28 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "24" };

            runProperties34.Append(runFonts37);
            runProperties34.Append(color28);
            runProperties34.Append(kern37);
            runProperties34.Append(fontSize28);
            runProperties34.Append(fontSizeComplexScript37);
            Text text28 = new Text();
            text28.Text = "中";

            run34.Append(runProperties34);
            run34.Append(text28);

            Run run35 = new Run();

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color29 = new Color() { Val = "000000" };
            Kern kern38 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize29 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "24" };

            runProperties35.Append(runFonts38);
            runProperties35.Append(color29);
            runProperties35.Append(kern38);
            runProperties35.Append(fontSize29);
            runProperties35.Append(fontSizeComplexScript38);
            Text text29 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text29.Text = "　";

            run35.Append(runProperties35);
            run35.Append(text29);

            Run run36 = new Run();

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color30 = new Color() { Val = "000000" };
            Kern kern39 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize30 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "24" };

            runProperties36.Append(runFonts39);
            runProperties36.Append(color30);
            runProperties36.Append(kern39);
            runProperties36.Append(fontSize30);
            runProperties36.Append(fontSizeComplexScript39);
            Text text30 = new Text();
            text30.Text = "華";

            run36.Append(runProperties36);
            run36.Append(text30);

            Run run37 = new Run();

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color31 = new Color() { Val = "000000" };
            Kern kern40 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize31 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "24" };

            runProperties37.Append(runFonts40);
            runProperties37.Append(color31);
            runProperties37.Append(kern40);
            runProperties37.Append(fontSize31);
            runProperties37.Append(fontSizeComplexScript40);
            Text text31 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text31.Text = "　";

            run37.Append(runProperties37);
            run37.Append(text31);

            Run run38 = new Run();

            RunProperties runProperties38 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color32 = new Color() { Val = "000000" };
            Kern kern41 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize32 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "24" };

            runProperties38.Append(runFonts41);
            runProperties38.Append(color32);
            runProperties38.Append(kern41);
            runProperties38.Append(fontSize32);
            runProperties38.Append(fontSizeComplexScript41);
            Text text32 = new Text();
            text32.Text = "民";

            run38.Append(runProperties38);
            run38.Append(text32);

            Run run39 = new Run();

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color33 = new Color() { Val = "000000" };
            Kern kern42 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize33 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "24" };

            runProperties39.Append(runFonts42);
            runProperties39.Append(color33);
            runProperties39.Append(kern42);
            runProperties39.Append(fontSize33);
            runProperties39.Append(fontSizeComplexScript42);
            Text text33 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text33.Text = "　";

            run39.Append(runProperties39);
            run39.Append(text33);

            Run run40 = new Run();

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color34 = new Color() { Val = "000000" };
            Kern kern43 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize34 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "24" };

            runProperties40.Append(runFonts43);
            runProperties40.Append(color34);
            runProperties40.Append(kern43);
            runProperties40.Append(fontSize34);
            runProperties40.Append(fontSizeComplexScript43);
            Text text34 = new Text();
            text34.Text = "國";

            run40.Append(runProperties40);
            run40.Append(text34);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run34);
            paragraph9.Append(run35);
            paragraph9.Append(run36);
            paragraph9.Append(run37);
            paragraph9.Append(run38);
            paragraph9.Append(run39);
            paragraph9.Append(run40);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            FrameProperties frameProperties10 = new FrameProperties() { Width = "4800", Height = (UInt32Value)3840U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "4201", Y = "4217", HeightType = HeightRuleValues.Exact };

            Tabs tabs10 = new Tabs();
            TabStop tabStop56 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop57 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop58 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop59 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop60 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop61 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop62 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop63 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop64 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop65 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop66 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop67 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop68 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };

            tabs10.Append(tabStop56);
            tabs10.Append(tabStop57);
            tabs10.Append(tabStop58);
            tabs10.Append(tabStop59);
            tabs10.Append(tabStop60);
            tabs10.Append(tabStop61);
            tabs10.Append(tabStop62);
            tabs10.Append(tabStop63);
            tabs10.Append(tabStop64);
            tabs10.Append(tabStop65);
            tabs10.Append(tabStop66);
            tabs10.Append(tabStop67);
            tabs10.Append(tabStop68);
            AutoSpaceDE autoSpaceDE10 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN10 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent10 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern44 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties10.Append(runFonts44);
            paragraphMarkRunProperties10.Append(kern44);
            paragraphMarkRunProperties10.Append(fontSizeComplexScript44);

            paragraphProperties10.Append(frameProperties10);
            paragraphProperties10.Append(tabs10);
            paragraphProperties10.Append(autoSpaceDE10);
            paragraphProperties10.Append(autoSpaceDN10);
            paragraphProperties10.Append(adjustRightIndent10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            Run run41 = new Run();

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color35 = new Color() { Val = "000000" };
            Kern kern45 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize35 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "24" };

            runProperties41.Append(runFonts45);
            runProperties41.Append(color35);
            runProperties41.Append(kern45);
            runProperties41.Append(fontSize35);
            runProperties41.Append(fontSizeComplexScript45);
            Text text35 = new Text();
            text35.Text = dt.Rows[i]["ppl"].ToString();

            run41.Append(runProperties41);

            string[] ppl = dt.Rows[i]["ppl"].ToString().Split(',');
            for (int p = 0; p < ppl.Length; p++)
            {
                run41.AppendChild(new Text(ppl[p]));
                run41.AppendChild(new Break());
            }

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run41);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            FrameProperties frameProperties11 = new FrameProperties() { Width = "4800", Height = (UInt32Value)3840U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "4201", Y = "4217", HeightType = HeightRuleValues.Exact };

            Tabs tabs11 = new Tabs();
            TabStop tabStop69 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop70 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop71 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop72 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop73 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop74 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop75 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop76 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop77 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop78 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop79 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop80 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop81 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };

            tabs11.Append(tabStop69);
            tabs11.Append(tabStop70);
            tabs11.Append(tabStop71);
            tabs11.Append(tabStop72);
            tabs11.Append(tabStop73);
            tabs11.Append(tabStop74);
            tabs11.Append(tabStop75);
            tabs11.Append(tabStop76);
            tabs11.Append(tabStop77);
            tabs11.Append(tabStop78);
            tabs11.Append(tabStop79);
            tabs11.Append(tabStop80);
            tabs11.Append(tabStop81);
            AutoSpaceDE autoSpaceDE11 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN11 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent11 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern47 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties11.Append(runFonts47);
            paragraphMarkRunProperties11.Append(kern47);
            paragraphMarkRunProperties11.Append(fontSizeComplexScript47);

            paragraphProperties11.Append(frameProperties11);
            paragraphProperties11.Append(tabs11);
            paragraphProperties11.Append(autoSpaceDE11);
            paragraphProperties11.Append(autoSpaceDN11);
            paragraphProperties11.Append(adjustRightIndent11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            paragraph11.Append(paragraphProperties11);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            FrameProperties frameProperties12 = new FrameProperties() { Width = "3720", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "3601", Y = "2777", HeightType = HeightRuleValues.Exact };

            Tabs tabs12 = new Tabs();
            TabStop tabStop82 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop83 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop84 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop85 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop86 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop87 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop88 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop89 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop90 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop91 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };

            tabs12.Append(tabStop82);
            tabs12.Append(tabStop83);
            tabs12.Append(tabStop84);
            tabs12.Append(tabStop85);
            tabs12.Append(tabStop86);
            tabs12.Append(tabStop87);
            tabs12.Append(tabStop88);
            tabs12.Append(tabStop89);
            tabs12.Append(tabStop90);
            tabs12.Append(tabStop91);
            AutoSpaceDE autoSpaceDE12 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN12 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent12 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern50 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties12.Append(runFonts50);
            paragraphMarkRunProperties12.Append(kern50);
            paragraphMarkRunProperties12.Append(fontSizeComplexScript50);

            paragraphProperties12.Append(frameProperties12);
            paragraphProperties12.Append(tabs12);
            paragraphProperties12.Append(autoSpaceDE12);
            paragraphProperties12.Append(autoSpaceDN12);
            paragraphProperties12.Append(adjustRightIndent12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run45 = new Run();

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color39 = new Color() { Val = "000000" };
            Kern kern51 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize39 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "24" };

            runProperties45.Append(runFonts51);
            runProperties45.Append(color39);
            runProperties45.Append(kern51);
            runProperties45.Append(fontSize39);
            runProperties45.Append(fontSizeComplexScript51);
            Text text39 = new Text();
            text39.Text = dt.Rows[i]["DeptName"].ToString();

            run45.Append(runProperties45);
            run45.Append(text39);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run45);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            FrameProperties frameProperties13 = new FrameProperties() { Width = "1800", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "8761", Y = "2777", HeightType = HeightRuleValues.Exact };

            Tabs tabs13 = new Tabs();
            TabStop tabStop92 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop93 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop94 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop95 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };

            tabs13.Append(tabStop92);
            tabs13.Append(tabStop93);
            tabs13.Append(tabStop94);
            tabs13.Append(tabStop95);
            AutoSpaceDE autoSpaceDE13 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN13 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent13 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern53 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties13.Append(runFonts53);
            paragraphMarkRunProperties13.Append(kern53);
            paragraphMarkRunProperties13.Append(fontSizeComplexScript53);

            paragraphProperties13.Append(frameProperties13);
            paragraphProperties13.Append(tabs13);
            paragraphProperties13.Append(autoSpaceDE13);
            paragraphProperties13.Append(autoSpaceDN13);
            paragraphProperties13.Append(adjustRightIndent13);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run47 = new Run();

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color41 = new Color() { Val = "000000" };
            Kern kern54 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize41 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "24" };

            runProperties47.Append(runFonts54);
            runProperties47.Append(color41);
            runProperties47.Append(kern54);
            runProperties47.Append(fontSize41);
            runProperties47.Append(fontSizeComplexScript54);
            Text text41 = new Text();
            text41.Text = dt.Rows[i]["UID"].ToString();

            run47.Append(runProperties47);
            run47.Append(text41);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run47);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            FrameProperties frameProperties14 = new FrameProperties() { Width = "1800", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "3841", Y = "13697", HeightType = HeightRuleValues.Exact };

            Tabs tabs14 = new Tabs();
            TabStop tabStop96 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop97 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop98 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop99 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };

            tabs14.Append(tabStop96);
            tabs14.Append(tabStop97);
            tabs14.Append(tabStop98);
            tabs14.Append(tabStop99);
            AutoSpaceDE autoSpaceDE14 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN14 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent14 = new AdjustRightIndent() { Val = false };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern56 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties14.Append(runFonts56);
            paragraphMarkRunProperties14.Append(kern56);
            paragraphMarkRunProperties14.Append(fontSizeComplexScript56);

            paragraphProperties14.Append(frameProperties14);
            paragraphProperties14.Append(tabs14);
            paragraphProperties14.Append(autoSpaceDE14);
            paragraphProperties14.Append(autoSpaceDN14);
            paragraphProperties14.Append(adjustRightIndent14);
            paragraphProperties14.Append(justification1);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run49 = new Run();

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color43 = new Color() { Val = "000000" };
            Kern kern57 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize43 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "24" };

            runProperties49.Append(runFonts57);
            runProperties49.Append(color43);
            runProperties49.Append(kern57);
            runProperties49.Append(fontSize43);
            runProperties49.Append(fontSizeComplexScript57);
            Text text43 = new Text();
            text43.Text = "一";

            run49.Append(runProperties49);
            run49.Append(text43);

            Run run50 = new Run();

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color44 = new Color() { Val = "000000" };
            Kern kern58 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize44 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "24" };

            runProperties50.Append(runFonts58);
            runProperties50.Append(color44);
            runProperties50.Append(kern58);
            runProperties50.Append(fontSize44);
            runProperties50.Append(fontSizeComplexScript58);
            Text text44 = new Text();
            string year = "";
            if (int.Parse(dt.Rows[i]["YEAR"].ToString()) - 100 < 10)
            {
                year += "○";
            }else {
                year += ChineseW[int.Parse(dt.Rows[i]["YEAR"].ToString().Substring(1, 1))];
            }
            year += ChineseW[int.Parse(dt.Rows[i]["YEAR"].ToString().Substring(2, 1))];
            
            text44.Text = year;
            
            run50.Append(runProperties50);
            run50.Append(text44);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run49);
            paragraph14.Append(run50);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            FrameProperties frameProperties15 = new FrameProperties() { Width = "1080", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "7921", Y = "13697", HeightType = HeightRuleValues.Exact };

            Tabs tabs15 = new Tabs();
            TabStop tabStop100 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop101 = new TabStop() { Val = TabStopValues.Left, Position = 720 };

            tabs15.Append(tabStop100);
            tabs15.Append(tabStop101);
            AutoSpaceDE autoSpaceDE15 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN15 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent15 = new AdjustRightIndent() { Val = false };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern60 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties15.Append(runFonts60);
            paragraphMarkRunProperties15.Append(kern60);
            paragraphMarkRunProperties15.Append(fontSizeComplexScript60);

            paragraphProperties15.Append(frameProperties15);
            paragraphProperties15.Append(tabs15);
            paragraphProperties15.Append(autoSpaceDE15);
            paragraphProperties15.Append(autoSpaceDN15);
            paragraphProperties15.Append(adjustRightIndent15);
            paragraphProperties15.Append(justification2);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run52 = new Run();

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts61 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color46 = new Color() { Val = "000000" };
            Kern kern61 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize46 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "24" };

            runProperties52.Append(runFonts61);
            runProperties52.Append(color46);
            runProperties52.Append(kern61);
            runProperties52.Append(fontSize46);
            runProperties52.Append(fontSizeComplexScript61);
            Text text46 = new Text();

            string day = "";
            if (int.Parse(dt.Rows[i]["DAY"].ToString()) <= 10)
            {
                day += ChineseW[int.Parse(dt.Rows[i]["DAY"].ToString())];
            }
            else if (int.Parse(dt.Rows[i]["DAY"].ToString()) > 10 && int.Parse(dt.Rows[i]["DAY"].ToString()) < 20)
            {
                day += ChineseW[10] + ChineseW[int.Parse(dt.Rows[i]["DAY"].ToString().Substring(1, 1))];
            }
            else if (int.Parse(dt.Rows[i]["DAY"].ToString()) == 20)
            {
                day += ChineseW[2] + ChineseW[10];
            }
            else if (int.Parse(dt.Rows[i]["DAY"].ToString()) > 20 && int.Parse(dt.Rows[i]["DAY"].ToString()) < 30)
            {
                day += ChineseW[2] + ChineseW[10] + ChineseW[int.Parse(dt.Rows[i]["DAY"].ToString().Substring(1, 1))];
            }
            else if (int.Parse(dt.Rows[i]["DAY"].ToString()) == 30)
            {
                day += ChineseW[3] + ChineseW[10];
            }
            else if (int.Parse(dt.Rows[i]["DAY"].ToString()) > 30)
            {
                day += ChineseW[3] + ChineseW[10] + ChineseW[int.Parse(dt.Rows[i]["DAY"].ToString().Substring(1, 1))];
            }

            text46.Text = day;

            run52.Append(runProperties52);
            run52.Append(text46);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run52);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            FrameProperties frameProperties16 = new FrameProperties() { Width = "1080", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "6241", Y = "13697", HeightType = HeightRuleValues.Exact };

            Tabs tabs16 = new Tabs();
            TabStop tabStop102 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop103 = new TabStop() { Val = TabStopValues.Left, Position = 720 };

            tabs16.Append(tabStop102);
            tabs16.Append(tabStop103);
            AutoSpaceDE autoSpaceDE16 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN16 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent16 = new AdjustRightIndent() { Val = false };
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern62 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties16.Append(runFonts62);
            paragraphMarkRunProperties16.Append(kern62);
            paragraphMarkRunProperties16.Append(fontSizeComplexScript62);

            paragraphProperties16.Append(frameProperties16);
            paragraphProperties16.Append(tabs16);
            paragraphProperties16.Append(autoSpaceDE16);
            paragraphProperties16.Append(autoSpaceDN16);
            paragraphProperties16.Append(adjustRightIndent16);
            paragraphProperties16.Append(justification3);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            Run run53 = new Run();

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color47 = new Color() { Val = "000000" };
            Kern kern63 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize47 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "24" };

            runProperties53.Append(runFonts63);
            runProperties53.Append(color47);
            runProperties53.Append(kern63);
            runProperties53.Append(fontSize47);
            runProperties53.Append(fontSizeComplexScript63);
            Text text47 = new Text();

            string month = "";
            if (int.Parse(dt.Rows[i]["MONTH"].ToString()) <= 10)
            {
                month += ChineseW[int.Parse(dt.Rows[i]["MONTH"].ToString())];
            }
            else if (int.Parse(dt.Rows[i]["MONTH"].ToString()) > 10)
            {
                month += ChineseW[10] + ChineseW[int.Parse(dt.Rows[i]["MONTH"].ToString().Substring(1, 1))];
            }
            text47.Text = month;

            run53.Append(runProperties53);
            run53.Append(text47);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run53);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            FrameProperties frameProperties17 = new FrameProperties() { Width = "360", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "5761", Y = "13697", HeightType = HeightRuleValues.Exact };
            AutoSpaceDE autoSpaceDE17 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN17 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent17 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern64 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties17.Append(runFonts64);
            paragraphMarkRunProperties17.Append(kern64);
            paragraphMarkRunProperties17.Append(fontSizeComplexScript64);

            paragraphProperties17.Append(frameProperties17);
            paragraphProperties17.Append(autoSpaceDE17);
            paragraphProperties17.Append(autoSpaceDN17);
            paragraphProperties17.Append(adjustRightIndent17);
            paragraphProperties17.Append(paragraphMarkRunProperties17);

            Run run54 = new Run();

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts65 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color48 = new Color() { Val = "000000" };
            Kern kern65 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize48 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "24" };

            runProperties54.Append(runFonts65);
            runProperties54.Append(color48);
            runProperties54.Append(kern65);
            runProperties54.Append(fontSize48);
            runProperties54.Append(fontSizeComplexScript65);
            Text text48 = new Text();
            text48.Text = "年";

            run54.Append(runProperties54);
            run54.Append(text48);

            paragraph17.Append(paragraphProperties17);
            paragraph17.Append(run54);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            FrameProperties frameProperties18 = new FrameProperties() { Width = "360", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "7441", Y = "13697", HeightType = HeightRuleValues.Exact };
            AutoSpaceDE autoSpaceDE18 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN18 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent18 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern66 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties18.Append(runFonts66);
            paragraphMarkRunProperties18.Append(kern66);
            paragraphMarkRunProperties18.Append(fontSizeComplexScript66);

            paragraphProperties18.Append(frameProperties18);
            paragraphProperties18.Append(autoSpaceDE18);
            paragraphProperties18.Append(autoSpaceDN18);
            paragraphProperties18.Append(adjustRightIndent18);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run55 = new Run();

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color49 = new Color() { Val = "000000" };
            Kern kern67 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize49 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "24" };

            runProperties55.Append(runFonts67);
            runProperties55.Append(color49);
            runProperties55.Append(kern67);
            runProperties55.Append(fontSize49);
            runProperties55.Append(fontSizeComplexScript67);
            Text text49 = new Text();
            text49.Text = "月";

            run55.Append(runProperties55);
            run55.Append(text49);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run55);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            FrameProperties frameProperties19 = new FrameProperties() { Width = "360", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "9121", Y = "13697", HeightType = HeightRuleValues.Exact };
            AutoSpaceDE autoSpaceDE19 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN19 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent19 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern68 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties19.Append(runFonts68);
            paragraphMarkRunProperties19.Append(kern68);
            paragraphMarkRunProperties19.Append(fontSizeComplexScript68);

            paragraphProperties19.Append(frameProperties19);
            paragraphProperties19.Append(autoSpaceDE19);
            paragraphProperties19.Append(autoSpaceDN19);
            paragraphProperties19.Append(adjustRightIndent19);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run56 = new Run();

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color50 = new Color() { Val = "000000" };
            Kern kern69 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize50 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "24" };

            runProperties56.Append(runFonts69);
            runProperties56.Append(color50);
            runProperties56.Append(kern69);
            runProperties56.Append(fontSize50);
            runProperties56.Append(fontSizeComplexScript69);
            Text text50 = new Text();
            text50.Text = "日";

            run56.Append(runProperties56);
            run56.Append(text50);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run56);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            FrameProperties frameProperties20 = new FrameProperties() { Width = "360", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "7321", Y = "2777", HeightType = HeightRuleValues.Exact };
            AutoSpaceDE autoSpaceDE20 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN20 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent20 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern70 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties20.Append(runFonts70);
            paragraphMarkRunProperties20.Append(kern70);
            paragraphMarkRunProperties20.Append(fontSizeComplexScript70);

            paragraphProperties20.Append(frameProperties20);
            paragraphProperties20.Append(autoSpaceDE20);
            paragraphProperties20.Append(autoSpaceDN20);
            paragraphProperties20.Append(adjustRightIndent20);
            paragraphProperties20.Append(paragraphMarkRunProperties20);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run57 = new Run();

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color51 = new Color() { Val = "000000" };
            Kern kern71 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize51 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "24" };

            runProperties57.Append(runFonts71);
            runProperties57.Append(color51);
            runProperties57.Append(kern71);
            runProperties57.Append(fontSize51);
            runProperties57.Append(fontSizeComplexScript71);
            Text text51 = new Text();
            text51.Text = "（";

            run57.Append(runProperties57);
            run57.Append(text51);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(proofError1);
            paragraph20.Append(run57);
            paragraph20.Append(proofError2);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            FrameProperties frameProperties21 = new FrameProperties() { Width = "480", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "7681", Y = "2777", HeightType = HeightRuleValues.Exact };

            Tabs tabs17 = new Tabs();
            TabStop tabStop104 = new TabStop() { Val = TabStopValues.Left, Position = 360 };

            tabs17.Append(tabStop104);
            AutoSpaceDE autoSpaceDE21 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN21 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent21 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern72 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties21.Append(runFonts72);
            paragraphMarkRunProperties21.Append(kern72);
            paragraphMarkRunProperties21.Append(fontSizeComplexScript72);

            paragraphProperties21.Append(frameProperties21);
            paragraphProperties21.Append(tabs17);
            paragraphProperties21.Append(autoSpaceDE21);
            paragraphProperties21.Append(autoSpaceDN21);
            paragraphProperties21.Append(adjustRightIndent21);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run58 = new Run();

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts73 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color52 = new Color() { Val = "000000" };
            Kern kern73 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize52 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "24" };

            runProperties58.Append(runFonts73);
            runProperties58.Append(color52);
            runProperties58.Append(kern73);
            runProperties58.Append(fontSize52);
            runProperties58.Append(fontSizeComplexScript73);
            Text text52 = new Text();
            text52.Text = "一";

            run58.Append(runProperties58);
            run58.Append(text52);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run58);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            FrameProperties frameProperties22 = new FrameProperties() { Width = "5000", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "4321", Y = "3497", HeightType = HeightRuleValues.Exact };

            Tabs tabs18 = new Tabs();
            TabStop tabStop105 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop106 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop107 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop108 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop109 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop110 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop111 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop112 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop113 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop114 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop115 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop116 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop117 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };

            tabs18.Append(tabStop105);
            tabs18.Append(tabStop106);
            tabs18.Append(tabStop107);
            tabs18.Append(tabStop108);
            tabs18.Append(tabStop109);
            tabs18.Append(tabStop110);
            tabs18.Append(tabStop111);
            tabs18.Append(tabStop112);
            tabs18.Append(tabStop113);
            tabs18.Append(tabStop114);
            tabs18.Append(tabStop115);
            tabs18.Append(tabStop116);
            tabs18.Append(tabStop117);
            AutoSpaceDE autoSpaceDE22 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN22 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent22 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts74 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern74 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties22.Append(runFonts74);
            paragraphMarkRunProperties22.Append(kern74);
            paragraphMarkRunProperties22.Append(fontSizeComplexScript74);

            paragraphProperties22.Append(frameProperties22);
            paragraphProperties22.Append(tabs18);
            paragraphProperties22.Append(autoSpaceDE22);
            paragraphProperties22.Append(autoSpaceDN22);
            paragraphProperties22.Append(adjustRightIndent22);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            Run run59 = new Run();

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts75 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color53 = new Color() { Val = "000000" };
            Kern kern75 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize53 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "24" };

            runProperties59.Append(runFonts75);
            runProperties59.Append(color53);
            runProperties59.Append(kern75);
            runProperties59.Append(fontSize53);
            runProperties59.Append(fontSizeComplexScript75);
            Text text53 = new Text();
            text53.Text = dt.Rows[i]["Class"].ToString();

            run59.Append(runProperties59);
            run59.Append(text53);


            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(run59);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            FrameProperties frameProperties23 = new FrameProperties() { Width = "9720", Height = (UInt32Value)645U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1201", Y = "14297", HeightType = HeightRuleValues.Exact };

            Tabs tabs19 = new Tabs();
            TabStop tabStop118 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop119 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop120 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop121 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop122 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop123 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop124 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop125 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop126 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop127 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop128 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop129 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop130 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };
            TabStop tabStop131 = new TabStop() { Val = TabStopValues.Left, Position = 5040 };
            TabStop tabStop132 = new TabStop() { Val = TabStopValues.Left, Position = 5400 };
            TabStop tabStop133 = new TabStop() { Val = TabStopValues.Left, Position = 5760 };
            TabStop tabStop134 = new TabStop() { Val = TabStopValues.Left, Position = 6120 };
            TabStop tabStop135 = new TabStop() { Val = TabStopValues.Left, Position = 6480 };
            TabStop tabStop136 = new TabStop() { Val = TabStopValues.Left, Position = 6840 };
            TabStop tabStop137 = new TabStop() { Val = TabStopValues.Left, Position = 7200 };
            TabStop tabStop138 = new TabStop() { Val = TabStopValues.Left, Position = 7560 };
            TabStop tabStop139 = new TabStop() { Val = TabStopValues.Left, Position = 7920 };
            TabStop tabStop140 = new TabStop() { Val = TabStopValues.Left, Position = 8280 };
            TabStop tabStop141 = new TabStop() { Val = TabStopValues.Left, Position = 8640 };
            TabStop tabStop142 = new TabStop() { Val = TabStopValues.Left, Position = 9000 };
            TabStop tabStop143 = new TabStop() { Val = TabStopValues.Left, Position = 9360 };

            tabs19.Append(tabStop118);
            tabs19.Append(tabStop119);
            tabs19.Append(tabStop120);
            tabs19.Append(tabStop121);
            tabs19.Append(tabStop122);
            tabs19.Append(tabStop123);
            tabs19.Append(tabStop124);
            tabs19.Append(tabStop125);
            tabs19.Append(tabStop126);
            tabs19.Append(tabStop127);
            tabs19.Append(tabStop128);
            tabs19.Append(tabStop129);
            tabs19.Append(tabStop130);
            tabs19.Append(tabStop131);
            tabs19.Append(tabStop132);
            tabs19.Append(tabStop133);
            tabs19.Append(tabStop134);
            tabs19.Append(tabStop135);
            tabs19.Append(tabStop136);
            tabs19.Append(tabStop137);
            tabs19.Append(tabStop138);
            tabs19.Append(tabStop139);
            tabs19.Append(tabStop140);
            tabs19.Append(tabStop141);
            tabs19.Append(tabStop142);
            tabs19.Append(tabStop143);
            AutoSpaceDE autoSpaceDE23 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN23 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent23 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern77 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties23.Append(runFonts77);
            paragraphMarkRunProperties23.Append(kern77);
            paragraphMarkRunProperties23.Append(fontSizeComplexScript77);

            paragraphProperties23.Append(frameProperties23);
            paragraphProperties23.Append(tabs19);
            paragraphProperties23.Append(autoSpaceDE23);
            paragraphProperties23.Append(autoSpaceDN23);
            paragraphProperties23.Append(adjustRightIndent23);
            paragraphProperties23.Append(paragraphMarkRunProperties23);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run61 = new Run();

            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts78 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color55 = new Color() { Val = "000000" };
            Kern kern78 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize55 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "24" };

            runProperties61.Append(runFonts78);
            runProperties61.Append(color55);
            runProperties61.Append(kern78);
            runProperties61.Append(fontSize55);
            runProperties61.Append(fontSizeComplexScript78);
            Text text55 = new Text();
            text55.Text = "註";

            run61.Append(runProperties61);
            run61.Append(text55);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run62 = new Run();

            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color56 = new Color() { Val = "000000" };
            Kern kern79 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize56 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "24" };

            runProperties62.Append(runFonts79);
            runProperties62.Append(color56);
            runProperties62.Append(kern79);
            runProperties62.Append(fontSize56);
            runProperties62.Append(fontSizeComplexScript79);
            Text text56 = new Text();
            text56.Text = "：";

            run62.Append(runProperties62);
            run62.Append(text56);

            Run run63 = new Run();

            RunProperties runProperties63 = new RunProperties();
            RunFonts runFonts80 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color57 = new Color() { Val = "000000" };
            Kern kern80 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize57 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "24" };

            runProperties63.Append(runFonts80);
            runProperties63.Append(color57);
            runProperties63.Append(kern80);
            runProperties63.Append(fontSize57);
            runProperties63.Append(fontSizeComplexScript80);
            Text text57 = new Text();
            text57.Text = "一";

            run63.Append(runProperties63);
            run63.Append(text57);

            Run run64 = new Run();

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color58 = new Color() { Val = "000000" };
            Kern kern81 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize58 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "24" };

            runProperties64.Append(runFonts81);
            runProperties64.Append(color58);
            runProperties64.Append(kern81);
            runProperties64.Append(fontSize58);
            runProperties64.Append(fontSizeComplexScript81);
            Text text58 = new Text();
            text58.Text = "、";

            run64.Append(runProperties64);
            run64.Append(text58);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run65 = new Run();

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts82 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color59 = new Color() { Val = "000000" };
            Kern kern82 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize59 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "24" };

            runProperties65.Append(runFonts82);
            runProperties65.Append(color59);
            runProperties65.Append(kern82);
            runProperties65.Append(fontSize59);
            runProperties65.Append(fontSizeComplexScript82);
            Text text59 = new Text();
            text59.Text = "本聯由";

            run65.Append(runProperties65);
            run65.Append(text59);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run66 = new Run();

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts83 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color60 = new Color() { Val = "000000" };
            Kern kern83 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize60 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "24" };

            runProperties66.Append(runFonts83);
            runProperties66.Append(color60);
            runProperties66.Append(kern83);
            runProperties66.Append(fontSize60);
            runProperties66.Append(fontSizeComplexScript83);
            Text text60 = new Text();
            text60.Text = "查核人員";

            run66.Append(runProperties66);
            run66.Append(text60);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run67 = new Run();

            RunProperties runProperties67 = new RunProperties();
            RunFonts runFonts84 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color61 = new Color() { Val = "000000" };
            Kern kern84 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize61 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "24" };

            runProperties67.Append(runFonts84);
            runProperties67.Append(color61);
            runProperties67.Append(kern84);
            runProperties67.Append(fontSize61);
            runProperties67.Append(fontSizeComplexScript84);
            Text text61 = new Text();
            text61.Text = "面交受檢";

            run67.Append(runProperties67);
            run67.Append(text61);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run68 = new Run();

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts85 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color62 = new Color() { Val = "000000" };
            Kern kern85 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize62 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "24" };

            runProperties68.Append(runFonts85);
            runProperties68.Append(color62);
            runProperties68.Append(kern85);
            runProperties68.Append(fontSize62);
            runProperties68.Append(fontSizeComplexScript85);
            Text text62 = new Text();
            text62.Text = "單位負責人收存";

            run68.Append(runProperties68);
            run68.Append(text62);

            Run run69 = new Run();

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color63 = new Color() { Val = "000000" };
            Kern kern86 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize63 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "24" };

            runProperties69.Append(runFonts86);
            runProperties69.Append(color63);
            runProperties69.Append(kern86);
            runProperties69.Append(fontSize63);
            runProperties69.Append(fontSizeComplexScript86);
            Text text63 = new Text();
            text63.Text = "。";

            run69.Append(runProperties69);
            run69.Append(text63);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(proofError3);
            paragraph23.Append(run61);
            paragraph23.Append(proofError4);
            paragraph23.Append(run62);
            paragraph23.Append(run63);
            paragraph23.Append(run64);
            paragraph23.Append(proofError5);
            paragraph23.Append(run65);
            paragraph23.Append(proofError6);
            paragraph23.Append(run66);
            paragraph23.Append(proofError7);
            paragraph23.Append(run67);
            paragraph23.Append(proofError8);
            paragraph23.Append(run68);
            paragraph23.Append(run69);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "00BB0351", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            FrameProperties frameProperties24 = new FrameProperties() { Width = "9720", Height = (UInt32Value)645U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1201", Y = "14297", HeightType = HeightRuleValues.Exact };

            Tabs tabs20 = new Tabs();
            TabStop tabStop144 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop145 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop146 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop147 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop148 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop149 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop150 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop151 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop152 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop153 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop154 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop155 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop156 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };
            TabStop tabStop157 = new TabStop() { Val = TabStopValues.Left, Position = 5040 };
            TabStop tabStop158 = new TabStop() { Val = TabStopValues.Left, Position = 5400 };
            TabStop tabStop159 = new TabStop() { Val = TabStopValues.Left, Position = 5760 };
            TabStop tabStop160 = new TabStop() { Val = TabStopValues.Left, Position = 6120 };
            TabStop tabStop161 = new TabStop() { Val = TabStopValues.Left, Position = 6480 };
            TabStop tabStop162 = new TabStop() { Val = TabStopValues.Left, Position = 6840 };
            TabStop tabStop163 = new TabStop() { Val = TabStopValues.Left, Position = 7200 };
            TabStop tabStop164 = new TabStop() { Val = TabStopValues.Left, Position = 7560 };
            TabStop tabStop165 = new TabStop() { Val = TabStopValues.Left, Position = 7920 };
            TabStop tabStop166 = new TabStop() { Val = TabStopValues.Left, Position = 8280 };
            TabStop tabStop167 = new TabStop() { Val = TabStopValues.Left, Position = 8640 };
            TabStop tabStop168 = new TabStop() { Val = TabStopValues.Left, Position = 9000 };
            TabStop tabStop169 = new TabStop() { Val = TabStopValues.Left, Position = 9360 };

            tabs20.Append(tabStop144);
            tabs20.Append(tabStop145);
            tabs20.Append(tabStop146);
            tabs20.Append(tabStop147);
            tabs20.Append(tabStop148);
            tabs20.Append(tabStop149);
            tabs20.Append(tabStop150);
            tabs20.Append(tabStop151);
            tabs20.Append(tabStop152);
            tabs20.Append(tabStop153);
            tabs20.Append(tabStop154);
            tabs20.Append(tabStop155);
            tabs20.Append(tabStop156);
            tabs20.Append(tabStop157);
            tabs20.Append(tabStop158);
            tabs20.Append(tabStop159);
            tabs20.Append(tabStop160);
            tabs20.Append(tabStop161);
            tabs20.Append(tabStop162);
            tabs20.Append(tabStop163);
            tabs20.Append(tabStop164);
            tabs20.Append(tabStop165);
            tabs20.Append(tabStop166);
            tabs20.Append(tabStop167);
            tabs20.Append(tabStop168);
            tabs20.Append(tabStop169);
            AutoSpaceDE autoSpaceDE24 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN24 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent24 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern87 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties24.Append(runFonts87);
            paragraphMarkRunProperties24.Append(kern87);
            paragraphMarkRunProperties24.Append(fontSizeComplexScript87);

            paragraphProperties24.Append(frameProperties24);
            paragraphProperties24.Append(tabs20);
            paragraphProperties24.Append(autoSpaceDE24);
            paragraphProperties24.Append(autoSpaceDN24);
            paragraphProperties24.Append(adjustRightIndent24);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            Run run70 = new Run();

            RunProperties runProperties70 = new RunProperties();
            RunFonts runFonts88 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color64 = new Color() { Val = "000000" };
            Kern kern88 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize64 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "24" };

            runProperties70.Append(runFonts88);
            runProperties70.Append(color64);
            runProperties70.Append(kern88);
            runProperties70.Append(fontSize64);
            runProperties70.Append(fontSizeComplexScript88);
            Text text64 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text64.Text = "　　";

            run70.Append(runProperties70);
            run70.Append(text64);

            Run run71 = new Run();

            RunProperties runProperties71 = new RunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color65 = new Color() { Val = "000000" };
            Kern kern89 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize65 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "24" };

            runProperties71.Append(runFonts89);
            runProperties71.Append(color65);
            runProperties71.Append(kern89);
            runProperties71.Append(fontSize65);
            runProperties71.Append(fontSizeComplexScript89);
            Text text65 = new Text();
            text65.Text = "二";

            run71.Append(runProperties71);
            run71.Append(text65);

            Run run72 = new Run();

            RunProperties runProperties72 = new RunProperties();
            RunFonts runFonts90 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color66 = new Color() { Val = "000000" };
            Kern kern90 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize66 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "24" };

            runProperties72.Append(runFonts90);
            runProperties72.Append(color66);
            runProperties72.Append(kern90);
            runProperties72.Append(fontSize66);
            runProperties72.Append(fontSizeComplexScript90);
            Text text66 = new Text();
            text66.Text = "、";

            run72.Append(runProperties72);
            run72.Append(text66);

            Run run73 = new Run();

            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts91 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color67 = new Color() { Val = "000000" };
            Kern kern91 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize67 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "24" };

            runProperties73.Append(runFonts91);
            runProperties73.Append(color67);
            runProperties73.Append(kern91);
            runProperties73.Append(fontSize67);
            runProperties73.Append(fontSizeComplexScript91);
            Text text67 = new Text();
            text67.Text = "受檢單位如認有需要證明查核人員身分時";

            run73.Append(runProperties73);
            run73.Append(text67);

            Run run74 = new Run();

            RunProperties runProperties74 = new RunProperties();
            RunFonts runFonts92 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color68 = new Color() { Val = "000000" };
            Kern kern92 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize68 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "24" };

            runProperties74.Append(runFonts92);
            runProperties74.Append(color68);
            runProperties74.Append(kern92);
            runProperties74.Append(fontSize68);
            runProperties74.Append(fontSizeComplexScript92);
            Text text68 = new Text();
            text68.Text = "，";

            run74.Append(runProperties74);
            run74.Append(text68);

            Run run75 = new Run();

            RunProperties runProperties75 = new RunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color69 = new Color() { Val = "000000" };
            Kern kern93 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize69 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "24" };

            runProperties75.Append(runFonts93);
            runProperties75.Append(color69);
            runProperties75.Append(kern93);
            runProperties75.Append(fontSize69);
            runProperties75.Append(fontSizeComplexScript93);
            Text text69 = new Text();
            text69.Text = "可向查核人員查看本行服務證等證件";

            run75.Append(runProperties75);
            run75.Append(text69);

            Run run76 = new Run();

            RunProperties runProperties76 = new RunProperties();
            RunFonts runFonts94 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color70 = new Color() { Val = "000000" };
            Kern kern94 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize70 = new FontSize() { Val = "23" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "24" };

            runProperties76.Append(runFonts94);
            runProperties76.Append(color70);
            runProperties76.Append(kern94);
            runProperties76.Append(fontSize70);
            runProperties76.Append(fontSizeComplexScript94);
            Text text70 = new Text();
            text70.Text = "。";

            run76.Append(runProperties76);
            run76.Append(text70);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(run70);
            paragraph24.Append(run71);
            paragraph24.Append(run72);
            paragraph24.Append(run73);
            paragraph24.Append(run74);
            paragraph24.Append(run75);
            paragraph24.Append(run76);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "006369CA", RsidRunAdditionDefault = "006369CA" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            WidowControl widowControl1 = new WidowControl();

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern95 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties25.Append(runFonts95);
            paragraphMarkRunProperties25.Append(kern95);
            paragraphMarkRunProperties25.Append(fontSizeComplexScript95);

            paragraphProperties25.Append(widowControl1);
            paragraphProperties25.Append(paragraphMarkRunProperties25);


            paragraph25.Append(paragraphProperties25);
            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "00BB0351", RsidParagraphAddition = "00BB0351", RsidRunAdditionDefault = "00BB0351" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE28 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN28 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent28 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern96 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties28.Append(runFonts96);
            paragraphMarkRunProperties28.Append(kern96);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript96);

            paragraphProperties28.Append(autoSpaceDE28);
            paragraphProperties28.Append(autoSpaceDN28);
            paragraphProperties28.Append(adjustRightIndent28);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            Run run77 = new Run();

            RunProperties runProperties77 = new RunProperties();
            NoProof noProof7 = new NoProof();

            runProperties77.Append(noProof7);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();

            Drawing drawing2 = new Drawing();

            Wp.Anchor anchor2 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251684864U, BehindDoc = true, Locked = false, LayoutInCell = false, AllowOverlap = true };
            Wp.SimplePosition simplePosition2 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition2 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Page };
            Wp.PositionOffset positionOffset3 = new Wp.PositionOffset();
            positionOffset3.Text = "5029200";

            horizontalPosition2.Append(positionOffset3);

            Wp.VerticalPosition verticalPosition2 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Page };
            Wp.PositionOffset positionOffset4 = new Wp.PositionOffset();
            positionOffset4.Text = "6868160";

            verticalPosition2.Append(positionOffset4);
            Wp.Extent extent2 = new Wp.Extent() { Cx = 895350L, Cy = 876300L };
            Wp.EffectExtent effectExtent2 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone2 = new Wp.WrapNone();
            Wp.DocProperties docProperties2 = new Wp.DocProperties() { Id = (UInt32Value)27U, Name = "圖片 27" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties2 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks2 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties2.Append(graphicFrameLocks2);

            A.Graphic graphic2 = new A.Graphic();
            graphic2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData2 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture7 = new Pic.Picture();
            picture7.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties2 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties2 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture 27" };

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties2 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks2 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties2.Append(pictureLocks2);

            nonVisualPictureProperties2.Append(nonVisualDrawingProperties2);
            nonVisualPictureProperties2.Append(nonVisualPictureDrawingProperties2);

            Pic.BlipFill blipFill2 = new Pic.BlipFill();
            A.Blip blip2 = new A.Blip() { Embed = "rId10" };
            A.SourceRectangle sourceRectangle2 = new A.SourceRectangle();

            A.Stretch stretch2 = new A.Stretch();
            A.FillRectangle fillRectangle2 = new A.FillRectangle();

            stretch2.Append(fillRectangle2);

            blipFill2.Append(blip2);
            blipFill2.Append(sourceRectangle2);
            blipFill2.Append(stretch2);

            Pic.ShapeProperties shapeProperties2 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 895350L, Cy = 876300L };

            transform2D2.Append(offset2);
            transform2D2.Append(extents2);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);
            A.NoFill noFill2 = new A.NoFill();

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(noFill2);

            picture7.Append(nonVisualPictureProperties2);
            picture7.Append(blipFill2);
            picture7.Append(shapeProperties2);

            graphicData2.Append(picture7);

            graphic2.Append(graphicData2);

            anchor2.Append(simplePosition2);
            anchor2.Append(horizontalPosition2);
            anchor2.Append(verticalPosition2);
            anchor2.Append(extent2);
            anchor2.Append(effectExtent2);
            anchor2.Append(wrapNone2);
            anchor2.Append(docProperties2);
            anchor2.Append(nonVisualGraphicFrameDrawingProperties2);
            anchor2.Append(graphic2);

            drawing2.Append(anchor2);

            run77.Append(runProperties77);
            run77.Append(lastRenderedPageBreak1);
            run77.Append(drawing2);

            paragraph28.Append(paragraphProperties28);
            paragraph28.Append(run77);

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(paragraph3);
            body1.Append(paragraph4);
            body1.Append(paragraph5);
            body1.Append(paragraph6);
            body1.Append(paragraph7);
            body1.Append(paragraph8);
            body1.Append(paragraph9);
            body1.Append(paragraph10);
            body1.Append(paragraph11);
            body1.Append(paragraph12);
            body1.Append(paragraph13);
            body1.Append(paragraph14);
            body1.Append(paragraph15);
            body1.Append(paragraph16);
            body1.Append(paragraph17);
            body1.Append(paragraph18);
            body1.Append(paragraph19);
            body1.Append(paragraph20);
            body1.Append(paragraph21);
            body1.Append(paragraph22);
            body1.Append(paragraph23);
            body1.Append(paragraph24);
            body1.Append(paragraph25);
            body1.Append(paragraph28);
            return body1;
        }
        

        // Generates content of fontTablePart1.
        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "新細明體" };
            AltName altName1 = new AltName() { Val = "PMingLiU" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020500000000000000" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "88" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "A00002FF", UnicodeSignature1 = "28CFFCFA", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00100001", CodePageSignature1 = "00000000" };

            font2.Append(altName1);
            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "標楷體" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "03000509000000000000" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "88" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Script };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "080E0000", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00100001", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);

            fontTablePart1.Fonts = fonts1;
        }

        // Generates content of documentSettingsPart1.
        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings();
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            EmbedSystemFonts embedSystemFonts1 = new EmbedSystemFonts();
            BordersDoNotSurroundHeader bordersDoNotSurroundHeader1 = new BordersDoNotSurroundHeader();
            BordersDoNotSurroundFooter bordersDoNotSurroundFooter1 = new BordersDoNotSurroundFooter();
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 720 };
            DrawingGridHorizontalSpacing drawingGridHorizontalSpacing1 = new DrawingGridHorizontalSpacing() { Val = "120" };
            DrawingGridVerticalSpacing drawingGridVerticalSpacing1 = new DrawingGridVerticalSpacing() { Val = "120" };
            DisplayHorizontalDrawingGrid displayHorizontalDrawingGrid1 = new DisplayHorizontalDrawingGrid() { Val = 0 };
            DisplayVerticalDrawingGrid displayVerticalDrawingGrid1 = new DisplayVerticalDrawingGrid() { Val = 3 };
            DoNotUseMarginsForDrawingGridOrigin doNotUseMarginsForDrawingGridOrigin1 = new DoNotUseMarginsForDrawingGridOrigin();
            DoNotShadeFormData doNotShadeFormData1 = new DoNotShadeFormData();
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.CompressPunctuation };
            DoNotValidateAgainstSchema doNotValidateAgainstSchema1 = new DoNotValidateAgainstSchema();
            DoNotDemarcateInvalidXml doNotDemarcateInvalidXml1 = new DoNotDemarcateInvalidXml();

            HeaderShapeDefaults headerShapeDefaults1 = new HeaderShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults1 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 6146 };

            headerShapeDefaults1.Append(shapeDefaults1);

            FootnoteDocumentWideProperties footnoteDocumentWideProperties1 = new FootnoteDocumentWideProperties();
            FootnoteSpecialReference footnoteSpecialReference1 = new FootnoteSpecialReference() { Id = -1 };
            FootnoteSpecialReference footnoteSpecialReference2 = new FootnoteSpecialReference() { Id = 0 };

            footnoteDocumentWideProperties1.Append(footnoteSpecialReference1);
            footnoteDocumentWideProperties1.Append(footnoteSpecialReference2);

            EndnoteDocumentWideProperties endnoteDocumentWideProperties1 = new EndnoteDocumentWideProperties();
            EndnoteSpecialReference endnoteSpecialReference1 = new EndnoteSpecialReference() { Id = -1 };
            EndnoteSpecialReference endnoteSpecialReference2 = new EndnoteSpecialReference() { Id = 0 };

            endnoteDocumentWideProperties1.Append(endnoteSpecialReference1);
            endnoteDocumentWideProperties1.Append(endnoteSpecialReference2);

            Compatibility compatibility1 = new Compatibility();
            SpaceForUnderline spaceForUnderline1 = new SpaceForUnderline();
            BalanceSingleByteDoubleByteWidth balanceSingleByteDoubleByteWidth1 = new BalanceSingleByteDoubleByteWidth();
            DoNotLeaveBackslashAlone doNotLeaveBackslashAlone1 = new DoNotLeaveBackslashAlone();
            UnderlineTrailingSpaces underlineTrailingSpaces1 = new UnderlineTrailingSpaces();
            DoNotExpandShiftReturn doNotExpandShiftReturn1 = new DoNotExpandShiftReturn();
            AdjustLineHeightInTable adjustLineHeightInTable1 = new AdjustLineHeightInTable();
            UseFarEastLayout useFarEastLayout1 = new UseFarEastLayout();

            compatibility1.Append(spaceForUnderline1);
            compatibility1.Append(balanceSingleByteDoubleByteWidth1);
            compatibility1.Append(doNotLeaveBackslashAlone1);
            compatibility1.Append(underlineTrailingSpaces1);
            compatibility1.Append(doNotExpandShiftReturn1);
            compatibility1.Append(adjustLineHeightInTable1);
            compatibility1.Append(useFarEastLayout1);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00BB0351" };
            Rsid rsid1 = new Rsid() { Val = "000F2B42" };
            Rsid rsid2 = new Rsid() { Val = "0031301F" };
            Rsid rsid3 = new Rsid() { Val = "006369CA" };
            Rsid rsid4 = new Rsid() { Val = "006D71C0" };
            Rsid rsid5 = new Rsid() { Val = "00BB0351" };
            Rsid rsid6 = new Rsid() { Val = "00C6135B" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Off };
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "en-US", EastAsia = "zh-TW" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };
            DoNotIncludeSubdocsInStats doNotIncludeSubdocsInStats1 = new DoNotIncludeSubdocsInStats();
            DoNotAutoCompressPictures doNotAutoCompressPictures1 = new DoNotAutoCompressPictures();

            ShapeDefaults shapeDefaults2 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 6146 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults2.Append(shapeDefaults3);
            shapeDefaults2.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };

            settings1.Append(zoom1);
            settings1.Append(embedSystemFonts1);
            settings1.Append(bordersDoNotSurroundHeader1);
            settings1.Append(bordersDoNotSurroundFooter1);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(drawingGridHorizontalSpacing1);
            settings1.Append(drawingGridVerticalSpacing1);
            settings1.Append(displayHorizontalDrawingGrid1);
            settings1.Append(displayVerticalDrawingGrid1);
            settings1.Append(doNotUseMarginsForDrawingGridOrigin1);
            settings1.Append(doNotShadeFormData1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(doNotValidateAgainstSchema1);
            settings1.Append(doNotDemarcateInvalidXml1);
            settings1.Append(headerShapeDefaults1);
            settings1.Append(footnoteDocumentWideProperties1);
            settings1.Append(endnoteDocumentWideProperties1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(doNotIncludeSubdocsInStats1);
            settings1.Append(doNotAutoCompressPictures1);
            settings1.Append(shapeDefaults2);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of imagePart1.
        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts187 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            Kern kern187 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize136 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript187 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts187);
            runPropertiesBaseStyle1.Append(kern187);
            runPropertiesBaseStyle1.Append(fontSize136);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript187);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1 };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 59, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid7 = new Rsid() { Val = "000F2B42" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl3 = new WidowControl() { Val = false };

            styleParagraphProperties1.Append(widowControl3);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid7);
            style1.Append(styleParagraphProperties1);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "a0", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);
            style2.Append(unhideWhenUsed1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "a1", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed2);
            style3.Append(primaryStyle2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "a2", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            Style style5 = new Style() { Type = StyleValues.Paragraph, StyleId = "a3" };
            StyleName styleName5 = new StyleName() { Val = "header" };
            BasedOn basedOn1 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "a4" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            Rsid rsid8 = new Rsid() { Val = "00BB0351" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs41 = new Tabs();
            TabStop tabStop339 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop340 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs41.Append(tabStop339);
            tabs41.Append(tabStop340);
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };

            styleParagraphProperties2.Append(tabs41);
            styleParagraphProperties2.Append(snapToGrid1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            FontSize fontSize137 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript188 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties1.Append(fontSize137);
            styleRunProperties1.Append(fontSizeComplexScript188);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(linkedStyle1);
            style5.Append(uIPriority4);
            style5.Append(semiHidden4);
            style5.Append(unhideWhenUsed4);
            style5.Append(rsid8);
            style5.Append(styleParagraphProperties2);
            style5.Append(styleRunProperties1);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName6 = new StyleName() { Val = "頁首 字元" };
            BasedOn basedOn2 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a3" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Rsid rsid9 = new Rsid() { Val = "00BB0351" };

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            FontSize fontSize138 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript189 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties2.Append(fontSize138);
            styleRunProperties2.Append(fontSizeComplexScript189);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle2);
            style6.Append(uIPriority5);
            style6.Append(semiHidden5);
            style6.Append(rsid9);
            style6.Append(styleRunProperties2);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5" };
            StyleName styleName7 = new StyleName() { Val = "footer" };
            BasedOn basedOn3 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "a6" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            Rsid rsid10 = new Rsid() { Val = "00BB0351" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs42 = new Tabs();
            TabStop tabStop341 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop342 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs42.Append(tabStop341);
            tabs42.Append(tabStop342);
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };

            styleParagraphProperties3.Append(tabs42);
            styleParagraphProperties3.Append(snapToGrid2);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            FontSize fontSize139 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript190 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties3.Append(fontSize139);
            styleRunProperties3.Append(fontSizeComplexScript190);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(linkedStyle3);
            style7.Append(uIPriority6);
            style7.Append(semiHidden6);
            style7.Append(unhideWhenUsed5);
            style7.Append(rsid10);
            style7.Append(styleParagraphProperties3);
            style7.Append(styleRunProperties3);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
            StyleName styleName8 = new StyleName() { Val = "頁尾 字元" };
            BasedOn basedOn4 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a5" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden7 = new SemiHidden();
            Rsid rsid11 = new Rsid() { Val = "00BB0351" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            FontSize fontSize140 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript191 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties4.Append(fontSize140);
            styleRunProperties4.Append(fontSizeComplexScript191);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(linkedStyle4);
            style8.Append(uIPriority7);
            style8.Append(semiHidden7);
            style8.Append(rsid11);
            style8.Append(styleRunProperties4);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);
            styles1.Append(style8);

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of customXmlPart1.
        private void GenerateCustomXmlPart1Content(CustomXmlPart customXmlPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<b:Sources SelectedStyle=\"\\APA.XSL\" StyleName=\"APA Fifth Edition\" xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\"></b:Sources>\r\n");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart1.
        private void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
        {
            Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem() { ItemId = "{FE0D128A-F09C-4C5F-9E4C-D0B34EFE7B26}" };
            dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference1 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography" };

            schemaReferences1.Append(schemaReference1);

            dataStoreItem1.Append(schemaReferences1);

            customXmlPropertiesPart1.DataStoreItem = dataStoreItem1;
        }

        // Generates content of endnotesPart1.
        private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
        {
            Endnotes endnotes1 = new Endnotes();
            endnotes1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph51 = new Paragraph() { RsidParagraphAddition = "00C6135B", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00C6135B" };

            Run run149 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run149.Append(separatorMark1);

            paragraph51.Append(run149);

            endnote1.Append(paragraph51);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph52 = new Paragraph() { RsidParagraphAddition = "00C6135B", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00C6135B" };

            Run run150 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run150.Append(continuationSeparatorMark1);

            paragraph52.Append(run150);

            endnote2.Append(paragraph52);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of footnotesPart1.
        private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
        {
            Footnotes footnotes1 = new Footnotes();
            footnotes1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

            Paragraph paragraph53 = new Paragraph() { RsidParagraphAddition = "00C6135B", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00C6135B" };

            Run run151 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run151.Append(separatorMark2);

            paragraph53.Append(run151);

            footnote1.Append(paragraph53);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph54 = new Paragraph() { RsidParagraphAddition = "00C6135B", RsidParagraphProperties = "00BB0351", RsidRunAdditionDefault = "00C6135B" };

            Run run152 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run152.Append(continuationSeparatorMark2);

            paragraph54.Append(run152);

            footnote2.Append(paragraph54);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of webSettingsPart1.
        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings();
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        // Generates content of imagePart2.
        private void GenerateImagePart2Content(ImagePart imagePart2)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart2Data);
            imagePart2.FeedData(data);
            data.Close();
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Office 佈景主題" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ????" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "?? ??" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "?? ??" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Crystal Reports";
            document.PackageProperties.Description = "Powered By Crystal";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-01-28T03:54:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-01-28T03:54:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Sony";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2014-01-28T03:39:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }

        #region Binary Data
        private string imagePart2Data = "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";

        private string imagePart1Data = "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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        #endregion
    }
}
