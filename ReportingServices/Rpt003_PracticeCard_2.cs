using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using System.Data;

namespace Rpt003_PracticeCard2
{
    public class GeneratedClass
    {
        // Data Source
        public DataTable dt { get; set; }

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

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId7");
            GenerateThemePart1Content(themePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId6");
            GenerateFontTablePart1Content(fontTablePart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId5");
            GenerateEndnotesPart1Content(endnotesPart1);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId4");
            GenerateFootnotesPart1Content(footnotesPart1);

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
            totalTime1.Text = "2";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "16";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "96";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "1";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "Crystal Decisions";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "111";
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00FD3847", RsidParagraphProperties = "00FD3847", RsidRunAdditionDefault = "00FD3847" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            FrameProperties frameProperties1 = new FrameProperties() { Width = "4816", Height = (UInt32Value)375U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1201", Y = "1726", HeightType = HeightRuleValues.Exact };

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            tabs1.Append(tabStop3);
            tabs1.Append(tabStop4);
            tabs1.Append(tabStop5);
            tabs1.Append(tabStop6);
            tabs1.Append(tabStop7);
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

            Run run1 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color1 = new Color() { Val = "000000" };
            Kern kern2 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize1 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "24" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(color1);
            runProperties1.Append(kern2);
            runProperties1.Append(fontSize1);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = "人力資源處台照";

            run1.Append(runProperties1);
            run1.Append(text1);

            Run run2 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "MS Mincho", HighAnsi = "MS Mincho", EastAsia = "MS Mincho", ComplexScript = "MS Mincho" };
            Color color2 = new Color() { Val = "000000" };
            Kern kern3 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize2 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

            runProperties2.Append(runFonts3);
            runProperties2.Append(color2);
            runProperties2.Append(kern3);
            runProperties2.Append(fontSize2);
            runProperties2.Append(fontSizeComplexScript3);
            Text text2 = new Text();
            text2.Text = "‎";

            run2.Append(runProperties2);
            run2.Append(text2);

            Run run3 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體", EastAsia = "標楷體", ComplexScript = "標楷體" };
            Color color3 = new Color() { Val = "000000" };
            Kern kern4 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize3 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

            runProperties3.Append(runFonts4);
            runProperties3.Append(color3);
            runProperties3.Append(kern4);
            runProperties3.Append(fontSize3);
            runProperties3.Append(fontSizeComplexScript4);
            Text text3 = new Text();
            text3.Text = "：";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);
            paragraph1.Append(run3);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00FD3847", RsidRunAdditionDefault = "00FD3847" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            FrameProperties frameProperties2 = new FrameProperties() { Width = "2040", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "6121", Y = "8086", HeightType = HeightRuleValues.Exact };

            Tabs tabs2 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };

            tabs2.Append(tabStop8);
            tabs2.Append(tabStop9);
            tabs2.Append(tabStop10);
            tabs2.Append(tabStop11);
            tabs2.Append(tabStop12);
            AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern5 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties2.Append(runFonts5);
            paragraphMarkRunProperties2.Append(kern5);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript5);

            paragraphProperties2.Append(frameProperties2);
            paragraphProperties2.Append(tabs2);
            paragraphProperties2.Append(autoSpaceDE2);
            paragraphProperties2.Append(autoSpaceDN2);
            paragraphProperties2.Append(adjustRightIndent2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run4 = new Run();

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color4 = new Color() { Val = "000000" };
            Kern kern6 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize4 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

            runProperties4.Append(runFonts6);
            runProperties4.Append(color4);
            runProperties4.Append(kern6);
            runProperties4.Append(fontSize4);
            runProperties4.Append(fontSizeComplexScript6);
            Text text4 = new Text();
            text4.Text = "董事會稽核處";

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run4);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00FD3847", RsidRunAdditionDefault = "00FD3847" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            FrameProperties frameProperties3 = new FrameProperties() { Width = "2720", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "5761", Y = "8686", HeightType = HeightRuleValues.Exact };

            Tabs tabs3 = new Tabs();
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };

            tabs3.Append(tabStop13);
            tabs3.Append(tabStop14);
            tabs3.Append(tabStop15);
            tabs3.Append(tabStop16);
            tabs3.Append(tabStop17);
            tabs3.Append(tabStop18);
            tabs3.Append(tabStop19);
            AutoSpaceDE autoSpaceDE3 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN3 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent3 = new AdjustRightIndent() { Val = false };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern7 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties3.Append(runFonts7);
            paragraphMarkRunProperties3.Append(kern7);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript7);

            paragraphProperties3.Append(frameProperties3);
            paragraphProperties3.Append(tabs3);
            paragraphProperties3.Append(autoSpaceDE3);
            paragraphProperties3.Append(autoSpaceDN3);
            paragraphProperties3.Append(adjustRightIndent3);
            paragraphProperties3.Append(justification1);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run5 = new Run();

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color5 = new Color() { Val = "000000" };
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize5 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

            runProperties5.Append(runFonts8);
            runProperties5.Append(color5);
            runProperties5.Append(kern8);
            runProperties5.Append(fontSize5);
            runProperties5.Append(fontSizeComplexScript8);
            Text text5 = new Text();
            text5.Text = dt.Rows[0]["Year"].ToString();

            run5.Append(runProperties5);
            run5.Append(text5);

            Run run6 = new Run();

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color6 = new Color() { Val = "000000" };
            Kern kern9 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize6 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

            runProperties6.Append(runFonts9);
            runProperties6.Append(color6);
            runProperties6.Append(kern9);
            runProperties6.Append(fontSize6);
            runProperties6.Append(fontSizeComplexScript9);
            Text text6 = new Text();
            text6.Text = "年";

            run6.Append(runProperties6);
            run6.Append(text6);

            Run run7 = new Run();

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color7 = new Color() { Val = "000000" };
            Kern kern10 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize7 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

            runProperties7.Append(runFonts10);
            runProperties7.Append(color7);
            runProperties7.Append(kern10);
            runProperties7.Append(fontSize7);
            runProperties7.Append(fontSizeComplexScript10);
            Text text7 = new Text();
            text7.Text = dt.Rows[0]["Month"].ToString();

            run7.Append(runProperties7);
            run7.Append(text7);

            Run run8 = new Run();

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color8 = new Color() { Val = "000000" };
            Kern kern11 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize8 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

            runProperties8.Append(runFonts11);
            runProperties8.Append(color8);
            runProperties8.Append(kern11);
            runProperties8.Append(fontSize8);
            runProperties8.Append(fontSizeComplexScript11);
            Text text8 = new Text();
            text8.Text = "月";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run();

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Color color9 = new Color() { Val = "000000" };
            Kern kern12 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize9 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

            runProperties9.Append(runFonts12);
            runProperties9.Append(color9);
            runProperties9.Append(kern12);
            runProperties9.Append(fontSize9);
            runProperties9.Append(fontSizeComplexScript12);
            Text text9 = new Text();
            text9.Text = dt.Rows[0]["Day"].ToString();

            run9.Append(runProperties9);
            run9.Append(text9);

            Run run10 = new Run();

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color10 = new Color() { Val = "000000" };
            Kern kern13 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize10 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

            runProperties10.Append(runFonts13);
            runProperties10.Append(color10);
            runProperties10.Append(kern13);
            runProperties10.Append(fontSize10);
            runProperties10.Append(fontSizeComplexScript13);
            Text text10 = new Text();
            text10.Text = "日";

            run10.Append(runProperties10);
            run10.Append(text10);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run5);
            paragraph3.Append(run6);
            paragraph3.Append(run7);
            paragraph3.Append(run8);
            paragraph3.Append(run9);
            paragraph3.Append(run10);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00163870", RsidParagraphProperties = "00163870", RsidRunAdditionDefault = "00FD3847" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            FrameProperties frameProperties4 = new FrameProperties() { Width = "8071", Height = (UInt32Value)4561U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1876", Y = "3031", HeightType = HeightRuleValues.Exact };

            Tabs tabs4 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop26 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop27 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop28 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop29 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop30 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop31 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop32 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };
            TabStop tabStop33 = new TabStop() { Val = TabStopValues.Left, Position = 5040 };
            TabStop tabStop34 = new TabStop() { Val = TabStopValues.Left, Position = 5400 };
            TabStop tabStop35 = new TabStop() { Val = TabStopValues.Left, Position = 5760 };
            TabStop tabStop36 = new TabStop() { Val = TabStopValues.Left, Position = 6120 };
            TabStop tabStop37 = new TabStop() { Val = TabStopValues.Left, Position = 6480 };
            TabStop tabStop38 = new TabStop() { Val = TabStopValues.Left, Position = 6840 };
            TabStop tabStop39 = new TabStop() { Val = TabStopValues.Left, Position = 7200 };
            TabStop tabStop40 = new TabStop() { Val = TabStopValues.Left, Position = 7560 };
            TabStop tabStop41 = new TabStop() { Val = TabStopValues.Left, Position = 7920 };
            TabStop tabStop42 = new TabStop() { Val = TabStopValues.Left, Position = 8280 };
            TabStop tabStop43 = new TabStop() { Val = TabStopValues.Left, Position = 8640 };
            TabStop tabStop44 = new TabStop() { Val = TabStopValues.Left, Position = 9000 };

            tabs4.Append(tabStop20);
            tabs4.Append(tabStop21);
            tabs4.Append(tabStop22);
            tabs4.Append(tabStop23);
            tabs4.Append(tabStop24);
            tabs4.Append(tabStop25);
            tabs4.Append(tabStop26);
            tabs4.Append(tabStop27);
            tabs4.Append(tabStop28);
            tabs4.Append(tabStop29);
            tabs4.Append(tabStop30);
            tabs4.Append(tabStop31);
            tabs4.Append(tabStop32);
            tabs4.Append(tabStop33);
            tabs4.Append(tabStop34);
            tabs4.Append(tabStop35);
            tabs4.Append(tabStop36);
            tabs4.Append(tabStop37);
            tabs4.Append(tabStop38);
            tabs4.Append(tabStop39);
            tabs4.Append(tabStop40);
            tabs4.Append(tabStop41);
            tabs4.Append(tabStop42);
            tabs4.Append(tabStop43);
            tabs4.Append(tabStop44);
            AutoSpaceDE autoSpaceDE4 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN4 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent4 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color11 = new Color() { Val = "000000" };
            Kern kern14 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize11 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties4.Append(runFonts14);
            paragraphMarkRunProperties4.Append(color11);
            paragraphMarkRunProperties4.Append(kern14);
            paragraphMarkRunProperties4.Append(fontSize11);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript14);

            paragraphProperties4.Append(frameProperties4);
            paragraphProperties4.Append(tabs4);
            paragraphProperties4.Append(autoSpaceDE4);
            paragraphProperties4.Append(autoSpaceDN4);
            paragraphProperties4.Append(adjustRightIndent4);
            paragraphProperties4.Append(spacingBetweenLines1);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run11 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color12 = new Color() { Val = "000000" };
            Kern kern15 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize12 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

            runProperties11.Append(runFonts15);
            runProperties11.Append(color12);
            runProperties11.Append(kern15);
            runProperties11.Append(fontSize12);
            runProperties11.Append(fontSizeComplexScript15);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = "　　貴處";

            run11.Append(runProperties11);
            run11.Append(text11);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run12 = new Run() { RsidRunProperties = "00163870", RsidRunAddition = "00163870" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color13 = new Color() { Val = "000000" };
            Kern kern16 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize13 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

            runProperties12.Append(runFonts16);
            runProperties12.Append(color13);
            runProperties12.Append(kern16);
            runProperties12.Append(fontSize13);
            runProperties12.Append(fontSizeComplexScript16);
            Text text12 = new Text();
            text12.Text = dt.Rows[0]["context"].ToString();

            run12.Append(runProperties12);
            run12.Append(text12);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run14 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color15 = new Color() { Val = "000000" };
            Kern kern18 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize15 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

            runProperties14.Append(runFonts18);
            runProperties14.Append(color15);
            runProperties14.Append(kern18);
            runProperties14.Append(fontSize15);
            runProperties14.Append(fontSizeComplexScript18);
            Text text14 = new Text();
            text14.Text = "函";

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run11);
            paragraph4.Append(proofError1);
            paragraph4.Append(run12);
            paragraph4.Append(proofError2);
            paragraph4.Append(run14);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00FD3847", RsidParagraphAddition = "00FD3847", RsidParagraphProperties = "00163870", RsidRunAdditionDefault = "00FD3847" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            FrameProperties frameProperties5 = new FrameProperties() { Width = "8071", Height = (UInt32Value)4561U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1876", Y = "3031", HeightType = HeightRuleValues.Exact };

            Tabs tabs5 = new Tabs();
            TabStop tabStop45 = new TabStop() { Val = TabStopValues.Left, Position = 360 };
            TabStop tabStop46 = new TabStop() { Val = TabStopValues.Left, Position = 720 };
            TabStop tabStop47 = new TabStop() { Val = TabStopValues.Left, Position = 1080 };
            TabStop tabStop48 = new TabStop() { Val = TabStopValues.Left, Position = 1440 };
            TabStop tabStop49 = new TabStop() { Val = TabStopValues.Left, Position = 1800 };
            TabStop tabStop50 = new TabStop() { Val = TabStopValues.Left, Position = 2160 };
            TabStop tabStop51 = new TabStop() { Val = TabStopValues.Left, Position = 2520 };
            TabStop tabStop52 = new TabStop() { Val = TabStopValues.Left, Position = 2880 };
            TabStop tabStop53 = new TabStop() { Val = TabStopValues.Left, Position = 3240 };
            TabStop tabStop54 = new TabStop() { Val = TabStopValues.Left, Position = 3600 };
            TabStop tabStop55 = new TabStop() { Val = TabStopValues.Left, Position = 3960 };
            TabStop tabStop56 = new TabStop() { Val = TabStopValues.Left, Position = 4320 };
            TabStop tabStop57 = new TabStop() { Val = TabStopValues.Left, Position = 4680 };
            TabStop tabStop58 = new TabStop() { Val = TabStopValues.Left, Position = 5040 };
            TabStop tabStop59 = new TabStop() { Val = TabStopValues.Left, Position = 5400 };
            TabStop tabStop60 = new TabStop() { Val = TabStopValues.Left, Position = 5760 };
            TabStop tabStop61 = new TabStop() { Val = TabStopValues.Left, Position = 6120 };
            TabStop tabStop62 = new TabStop() { Val = TabStopValues.Left, Position = 6480 };
            TabStop tabStop63 = new TabStop() { Val = TabStopValues.Left, Position = 6840 };
            TabStop tabStop64 = new TabStop() { Val = TabStopValues.Left, Position = 7200 };
            TabStop tabStop65 = new TabStop() { Val = TabStopValues.Left, Position = 7560 };
            TabStop tabStop66 = new TabStop() { Val = TabStopValues.Left, Position = 7920 };
            TabStop tabStop67 = new TabStop() { Val = TabStopValues.Left, Position = 8280 };
            TabStop tabStop68 = new TabStop() { Val = TabStopValues.Left, Position = 8640 };
            TabStop tabStop69 = new TabStop() { Val = TabStopValues.Left, Position = 9000 };

            tabs5.Append(tabStop45);
            tabs5.Append(tabStop46);
            tabs5.Append(tabStop47);
            tabs5.Append(tabStop48);
            tabs5.Append(tabStop49);
            tabs5.Append(tabStop50);
            tabs5.Append(tabStop51);
            tabs5.Append(tabStop52);
            tabs5.Append(tabStop53);
            tabs5.Append(tabStop54);
            tabs5.Append(tabStop55);
            tabs5.Append(tabStop56);
            tabs5.Append(tabStop57);
            tabs5.Append(tabStop58);
            tabs5.Append(tabStop59);
            tabs5.Append(tabStop60);
            tabs5.Append(tabStop61);
            tabs5.Append(tabStop62);
            tabs5.Append(tabStop63);
            tabs5.Append(tabStop64);
            tabs5.Append(tabStop65);
            tabs5.Append(tabStop66);
            tabs5.Append(tabStop67);
            tabs5.Append(tabStop68);
            tabs5.Append(tabStop69);
            AutoSpaceDE autoSpaceDE5 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN5 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent5 = new AdjustRightIndent() { Val = false };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern19 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties5.Append(runFonts19);
            paragraphMarkRunProperties5.Append(kern19);
            paragraphMarkRunProperties5.Append(fontSizeComplexScript19);

            paragraphProperties5.Append(frameProperties5);
            paragraphProperties5.Append(tabs5);
            paragraphProperties5.Append(autoSpaceDE5);
            paragraphProperties5.Append(autoSpaceDN5);
            paragraphProperties5.Append(adjustRightIndent5);
            paragraphProperties5.Append(spacingBetweenLines2);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run15 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color16 = new Color() { Val = "000000" };
            Kern kern20 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize16 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

            runProperties15.Append(runFonts20);
            runProperties15.Append(color16);
            runProperties15.Append(kern20);
            runProperties15.Append(fontSize16);
            runProperties15.Append(fontSizeComplexScript20);
            Text text15 = new Text();
            text15.Text = "派查核實習人員";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunAddition = "00163870" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color17 = new Color() { Val = "000000" };
            Kern kern21 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize17 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "24" };

            runProperties16.Append(runFonts21);
            runProperties16.Append(color17);
            runProperties16.Append(kern21);
            runProperties16.Append(fontSize17);
            runProperties16.Append(fontSizeComplexScript21);
            Text text16 = new Text();
            text16.Text = dt.Rows[0]["ppl"].ToString();

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color18 = new Color() { Val = "000000" };
            Kern kern22 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize18 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

            runProperties17.Append(runFonts22);
            runProperties17.Append(color18);
            runProperties17.Append(kern22);
            runProperties17.Append(fontSize18);
            runProperties17.Append(fontSizeComplexScript22);
            Text text17 = new Text();
            text17.Text = "等";

            run17.Append(runProperties17);
            run17.Append(text17);

            Run run18 = new Run() { RsidRunAddition = "00163870" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color19 = new Color() { Val = "000000" };
            Kern kern23 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize19 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

            runProperties18.Append(runFonts23);
            runProperties18.Append(color19);
            runProperties18.Append(kern23);
            runProperties18.Append(fontSize19);
            runProperties18.Append(fontSizeComplexScript23);
            Text text18 = new Text();
            text18.Text = dt.Rows[0]["p_num"].ToString();

            run18.Append(runProperties18);
            run18.Append(text18);

            Run run19 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts24 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color20 = new Color() { Val = "000000" };
            Kern kern24 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize20 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };

            runProperties19.Append(runFonts24);
            runProperties19.Append(color20);
            runProperties19.Append(kern24);
            runProperties19.Append(fontSize20);
            runProperties19.Append(fontSizeComplexScript24);
            Text text19 = new Text();
            text19.Text = "名，";

            run19.Append(runProperties19);
            run19.Append(text19);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run20 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color21 = new Color() { Val = "000000" };
            Kern kern25 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize21 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

            runProperties20.Append(runFonts25);
            runProperties20.Append(color21);
            runProperties20.Append(kern25);
            runProperties20.Append(fontSize21);
            runProperties20.Append(fontSizeComplexScript25);
            Text text20 = new Text();
            text20.Text = "均已依";

            run20.Append(runProperties20);
            run20.Append(text20);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run21 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color22 = new Color() { Val = "000000" };
            Kern kern26 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize22 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

            runProperties21.Append(runFonts26);
            runProperties21.Append(color22);
            runProperties21.Append(kern26);
            runProperties21.Append(fontSize22);
            runProperties21.Append(fontSizeComplexScript26);
            Text text21 = new Text();
            text21.Text = "計畫實習期滿，並撰寫查核實習心得報告經核";

            run21.Append(runProperties21);
            run21.Append(text21);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run22 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color23 = new Color() { Val = "000000" };
            Kern kern27 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize23 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

            runProperties22.Append(runFonts27);
            runProperties22.Append(color23);
            runProperties22.Append(kern27);
            runProperties22.Append(fontSize23);
            runProperties22.Append(fontSizeComplexScript27);
            Text text22 = new Text();
            text22.Text = "可留卷備查";

            run22.Append(runProperties22);
            run22.Append(text22);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run23 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color24 = new Color() { Val = "000000" };
            Kern kern28 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize24 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

            runProperties23.Append(runFonts28);
            runProperties23.Append(color24);
            runProperties23.Append(kern28);
            runProperties23.Append(fontSize24);
            runProperties23.Append(fontSizeComplexScript28);
            Text text23 = new Text();
            text23.Text = "，依規定核發「查核實習證明書」共";

            run23.Append(runProperties23);
            run23.Append(text23);

            Run run24 = new Run() { RsidRunAddition = "00163870" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color25 = new Color() { Val = "000000" };
            Kern kern29 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize25 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "24" };

            runProperties24.Append(runFonts29);
            runProperties24.Append(color25);
            runProperties24.Append(kern29);
            runProperties24.Append(fontSize25);
            runProperties24.Append(fontSizeComplexScript29);
            Text text24 = new Text();
            text24.Text = dt.Rows[0]["p_num"].ToString();

            run24.Append(runProperties24);
            run24.Append(text24);

            Run run25 = new Run() { RsidRunProperties = "00FD3847" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
            Color color26 = new Color() { Val = "000000" };
            Kern kern30 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize26 = new FontSize() { Val = "31" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

            runProperties25.Append(runFonts30);
            runProperties25.Append(color26);
            runProperties25.Append(kern30);
            runProperties25.Append(fontSize26);
            runProperties25.Append(fontSizeComplexScript30);
            Text text25 = new Text();
            text25.Text = "份，請　查照。";

            run25.Append(runProperties25);
            run25.Append(text25);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run15);
            paragraph5.Append(run16);
            paragraph5.Append(run17);
            paragraph5.Append(run18);
            paragraph5.Append(run19);
            paragraph5.Append(proofError3);
            paragraph5.Append(run20);
            paragraph5.Append(proofError4);
            paragraph5.Append(run21);
            paragraph5.Append(proofError5);
            paragraph5.Append(run22);
            paragraph5.Append(proofError6);
            paragraph5.Append(run23);
            paragraph5.Append(run24);
            paragraph5.Append(run25);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00FD3847", RsidRunAdditionDefault = "00FD3847" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            AutoSpaceDE autoSpaceDE6 = new AutoSpaceDE() { Val = false };
            AutoSpaceDN autoSpaceDN6 = new AutoSpaceDN() { Val = false };
            AdjustRightIndent adjustRightIndent6 = new AdjustRightIndent() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
            Kern kern31 = new Kern() { Val = (UInt32Value)0U };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

            paragraphMarkRunProperties6.Append(runFonts31);
            paragraphMarkRunProperties6.Append(kern31);
            paragraphMarkRunProperties6.Append(fontSizeComplexScript31);

            paragraphProperties6.Append(autoSpaceDE6);
            paragraphProperties6.Append(autoSpaceDN6);
            paragraphProperties6.Append(adjustRightIndent6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            paragraph6.Append(paragraphProperties6);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "00FD3847" };
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

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(paragraph3);
            body1.Append(paragraph4);
            body1.Append(paragraph5);
            body1.Append(paragraph6);
            body1.Append(sectionProperties1);

            document1.Append(body1);

            mainDocumentPart1.Document = document1;
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
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
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
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
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
            Ovml.ShapeDefaults shapeDefaults1 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 3074 };

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
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00FD3847" };
            Rsid rsid1 = new Rsid() { Val = "00163870" };
            Rsid rsid2 = new Rsid() { Val = "00202D19" };
            Rsid rsid3 = new Rsid() { Val = "007A1762" };
            Rsid rsid4 = new Rsid() { Val = "00FD3847" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);

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
            Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 3074 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "2" };

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

        // Generates content of styleDefinitionsPart1.
        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts32 = new RunFonts() { ComplexScript = "Times New Roman", AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia };
            Kern kern32 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize27 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "24" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts32);
            runPropertiesBaseStyle1.Append(kern32);
            runPropertiesBaseStyle1.Append(fontSize27);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript32);
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

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl() { Val = false };

            styleParagraphProperties1.Append(widowControl1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts33 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties1.Append(runFonts33);
            styleRunProperties1.Append(fontSizeComplexScript33);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(styleParagraphProperties1);
            style1.Append(styleRunProperties1);

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
            Rsid rsid5 = new Rsid() { Val = "00FD3847" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs6 = new Tabs();
            TabStop tabStop70 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop71 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs6.Append(tabStop70);
            tabs6.Append(tabStop71);
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };

            styleParagraphProperties2.Append(tabs6);
            styleParagraphProperties2.Append(snapToGrid1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            FontSize fontSize28 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties2.Append(fontSize28);
            styleRunProperties2.Append(fontSizeComplexScript34);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(linkedStyle1);
            style5.Append(uIPriority4);
            style5.Append(semiHidden4);
            style5.Append(unhideWhenUsed4);
            style5.Append(rsid5);
            style5.Append(styleParagraphProperties2);
            style5.Append(styleRunProperties2);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName6 = new StyleName() { Val = "頁首 字元" };
            BasedOn basedOn2 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a3" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Locked locked1 = new Locked();
            Rsid rsid6 = new Rsid() { Val = "00FD3847" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            RunFonts runFonts34 = new RunFonts() { ComplexScript = "Times New Roman" };
            FontSize fontSize29 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties3.Append(runFonts34);
            styleRunProperties3.Append(fontSize29);
            styleRunProperties3.Append(fontSizeComplexScript35);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle2);
            style6.Append(uIPriority5);
            style6.Append(semiHidden5);
            style6.Append(locked1);
            style6.Append(rsid6);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5" };
            StyleName styleName7 = new StyleName() { Val = "footer" };
            BasedOn basedOn3 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "a6" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            Rsid rsid7 = new Rsid() { Val = "00FD3847" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs7 = new Tabs();
            TabStop tabStop72 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop73 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs7.Append(tabStop72);
            tabs7.Append(tabStop73);
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };

            styleParagraphProperties3.Append(tabs7);
            styleParagraphProperties3.Append(snapToGrid2);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            FontSize fontSize30 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties4.Append(fontSize30);
            styleRunProperties4.Append(fontSizeComplexScript36);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(linkedStyle3);
            style7.Append(uIPriority6);
            style7.Append(semiHidden6);
            style7.Append(unhideWhenUsed5);
            style7.Append(rsid7);
            style7.Append(styleParagraphProperties3);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
            StyleName styleName8 = new StyleName() { Val = "頁尾 字元" };
            BasedOn basedOn4 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a5" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden7 = new SemiHidden();
            Locked locked2 = new Locked();
            Rsid rsid8 = new Rsid() { Val = "00FD3847" };

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            RunFonts runFonts35 = new RunFonts() { ComplexScript = "Times New Roman" };
            FontSize fontSize31 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties5.Append(runFonts35);
            styleRunProperties5.Append(fontSize31);
            styleRunProperties5.Append(fontSizeComplexScript37);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(linkedStyle4);
            style8.Append(uIPriority7);
            style8.Append(semiHidden7);
            style8.Append(locked2);
            style8.Append(rsid8);
            style8.Append(styleRunProperties5);

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

            Font font5 = new Font() { Name = "MS Mincho" };
            AltName altName2 = new AltName() { Val = "?? ?玃" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "02020609040205080304" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "80" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Modern };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "6AC7FDFB", UnicodeSignature2 = "00000012", UnicodeSignature3 = "00000000", CodePageSignature0 = "0002009F", CodePageSignature1 = "00000000" };

            font5.Append(altName2);
            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Cambria" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "02040503050406030204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "400004FF", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font6.Append(panose1Number6);
            font6.Append(fontCharSet6);
            font6.Append(fontFamily6);
            font6.Append(pitch6);
            font6.Append(fontSignature6);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);

            fontTablePart1.Fonts = fonts1;
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

            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00202D19", RsidParagraphProperties = "00FD3847", RsidRunAdditionDefault = "00202D19" };

            Run run26 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run26.Append(separatorMark1);

            paragraph7.Append(run26);

            endnote1.Append(paragraph7);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00202D19", RsidParagraphProperties = "00FD3847", RsidRunAdditionDefault = "00202D19" };

            Run run27 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run27.Append(continuationSeparatorMark1);

            paragraph8.Append(run27);

            endnote2.Append(paragraph8);

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

            Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "00202D19", RsidParagraphProperties = "00FD3847", RsidRunAdditionDefault = "00202D19" };

            Run run28 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run28.Append(separatorMark2);

            paragraph9.Append(run28);

            footnote1.Append(paragraph9);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "00202D19", RsidParagraphProperties = "00FD3847", RsidRunAdditionDefault = "00202D19" };

            Run run29 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run29.Append(continuationSeparatorMark2);

            paragraph10.Append(run29);

            footnote2.Append(paragraph10);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Crystal Reports";
            document.PackageProperties.Description = "Powered By Crystal";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-02-12T05:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-02-12T05:40:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Sony";
        }


    }

}