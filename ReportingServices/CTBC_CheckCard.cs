using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using System.Data;
using System;

namespace CheckCard_CTBC
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

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId8");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId13");
            GenerateThemePart1Content(themePart1);

            CustomXmlPart customXmlPart1 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId3");
            GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId7");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId12");
            GenerateFontTablePart1Content(fontTablePart1);

            CustomXmlPart customXmlPart2 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId2");
            GenerateCustomXmlPart2Content(customXmlPart2);

            CustomXmlPropertiesPart customXmlPropertiesPart2 = customXmlPart2.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart2Content(customXmlPropertiesPart2);

            CustomXmlPart customXmlPart3 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId1");
            GenerateCustomXmlPart3Content(customXmlPart3);

            CustomXmlPropertiesPart customXmlPropertiesPart3 = customXmlPart3.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart3Content(customXmlPropertiesPart3);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId6");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            HeaderPart headerPart1 = mainDocumentPart1.AddNewPart<HeaderPart>("rId11");
            GenerateHeaderPart1Content(headerPart1);

            NumberingDefinitionsPart numberingDefinitionsPart1 = mainDocumentPart1.AddNewPart<NumberingDefinitionsPart>("rId5");
            GenerateNumberingDefinitionsPart1Content(numberingDefinitionsPart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId10");
            GenerateEndnotesPart1Content(endnotesPart1);

            CustomXmlPart customXmlPart4 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId4");
            GenerateCustomXmlPart4Content(customXmlPart4);

            CustomXmlPropertiesPart customXmlPropertiesPart4 = customXmlPart4.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart4Content(customXmlPropertiesPart4);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId9");
            GenerateFootnotesPart1Content(footnotesPart1);

            CustomFilePropertiesPart customFilePropertiesPart1 = document.AddNewPart<CustomFilePropertiesPart>("rId4");
            GenerateCustomFilePropertiesPart1Content(customFilePropertiesPart1);

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
            totalTime1.Text = "0";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "33";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "189";
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
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "221";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0000";

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
            Document document1 = new Document() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            document1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            document1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            document1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            document1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            document1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            document1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            document1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            document1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            document1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            document1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            document1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            document1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            document1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            document1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            document1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            document1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            document1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            document1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            document1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            document1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            document1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            foreach (DataRow rows in dt.Rows)
            {
                Body body1 = new Body();

                Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "009E00FF", RsidParagraphAddition = "00D66A62", RsidParagraphProperties = "00D66A62", RsidRunAdditionDefault = "00D66A62", ParagraphId = "2C6190AE", TextId = "77777777" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
                RunFonts runFonts1 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體", EastAsia = "標楷體" };

                paragraphMarkRunProperties1.Append(runFonts1);

                paragraphProperties1.Append(justification1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);

                paragraph1.Append(paragraphProperties1);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00D66A62", RsidParagraphProperties = "4CC70866", RsidRunAdditionDefault = "00D66A62", ParagraphId = "675720C6", TextId = "00CA4553" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                Justification justification2 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                RunFonts runFonts2 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color1 = new Color() { Val = "0000FF" };

                paragraphMarkRunProperties2.Append(runFonts2);
                paragraphMarkRunProperties2.Append(color1);

                paragraphProperties2.Append(justification2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run1 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties1 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Languages languages1 = new Languages() { EastAsia = "zh-HK" };

                runProperties1.Append(runFonts3);
                runProperties1.Append(languages1);
                Text text1 = new Text();
                text1.Text = "列印日期";

                run1.Append(runProperties1);
                run1.Append(text1);

                Run run2 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts4 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

                runProperties2.Append(runFonts4);
                Text text2 = new Text();
                text2.Text = "：";

                run2.Append(runProperties2);
                run2.Append(text2);

                Run run3 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts5 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color2 = new Color() { Val = "0000FF" };

                runProperties3.Append(runFonts5);
                runProperties3.Append(color2);
                Text text3 = new Text();
                text3.Text = "[";

                run3.Append(runProperties3);
                run3.Append(text3);

                Run run4 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties4 = new RunProperties();
                RunFonts runFonts6 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color3 = new Color() { Val = "0000FF" };
                Languages languages2 = new Languages() { EastAsia = "zh-HK" };

                runProperties4.Append(runFonts6);
                runProperties4.Append(color3);
                runProperties4.Append(languages2);
                Text text4 = new Text();
                text4.Text = Convert.ToDateTime(rows["printdate"]).ToString("yyyy-MM-dd");

                run4.Append(runProperties4);
                run4.Append(text4);

                Run run5 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties5 = new RunProperties();
                RunFonts runFonts7 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color4 = new Color() { Val = "0000FF" };

                runProperties5.Append(runFonts7);
                runProperties5.Append(color4);
                Text text5 = new Text();
                text5.Text = "]";

                run5.Append(runProperties5);
                run5.Append(text5);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run1);
                paragraph2.Append(run2);
                paragraph2.Append(run3);
                paragraph2.Append(run4);
                paragraph2.Append(run5);

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00D66A62", RsidParagraphProperties = "0007635F", RsidRunAdditionDefault = "00D66A62", ParagraphId = "13512E15", TextId = "77777777" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                Justification justification3 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                RunFonts runFonts8 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Bold bold1 = new Bold();
                FontSize fontSize1 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties3.Append(runFonts8);
                paragraphMarkRunProperties3.Append(bold1);
                paragraphMarkRunProperties3.Append(fontSize1);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript1);

                paragraphProperties3.Append(justification3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                paragraph3.Append(paragraphProperties3);

                Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00397F67", RsidParagraphAddition = "00CA3D7A", RsidParagraphProperties = "0007635F", RsidRunAdditionDefault = "0007635F", ParagraphId = "508D31C9", TextId = "4206A0BC" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                Justification justification4 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                RunFonts runFonts9 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Bold bold2 = new Bold();
                FontSize fontSize2 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties4.Append(runFonts9);
                paragraphMarkRunProperties4.Append(bold2);
                paragraphMarkRunProperties4.Append(fontSize2);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript2);

                paragraphProperties4.Append(justification4);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run6 = new Run() { RsidRunProperties = "00397F67" };

                RunProperties runProperties6 = new RunProperties();
                RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Bold bold3 = new Bold();
                FontSize fontSize3 = new FontSize() { Val = "36" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

                runProperties6.Append(runFonts10);
                runProperties6.Append(bold3);
                runProperties6.Append(fontSize3);
                runProperties6.Append(fontSizeComplexScript3);
                Text text6 = new Text();
                text6.Text = "中國信託商業銀行內部稽核檢查通知";

                run6.Append(runProperties6);
                run6.Append(text6);

                paragraph4.Append(paragraphProperties4);
                paragraph4.Append(run6);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "0007635F", RsidParagraphProperties = "0050595B", RsidRunAdditionDefault = "0007635F", ParagraphId = "4ABC6FB0", TextId = "77777777" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                RunFonts runFonts11 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize4 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };
                Languages languages3 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties5.Append(runFonts11);
                paragraphMarkRunProperties5.Append(fontSize4);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript4);
                paragraphMarkRunProperties5.Append(languages3);

                paragraphProperties5.Append(spacingBetweenLines1);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                paragraph5.Append(paragraphProperties5);

                Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00CA3D7A", RsidParagraphProperties = "0050595B", RsidRunAdditionDefault = "00A02E9E", ParagraphId = "53BBD3F2", TextId = "3F2A0988" };

                ParagraphProperties paragraphProperties6 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
                RunFonts runFonts12 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize5 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties6.Append(runFonts12);
                paragraphMarkRunProperties6.Append(fontSize5);
                paragraphMarkRunProperties6.Append(fontSizeComplexScript5);

                paragraphProperties6.Append(spacingBetweenLines2);
                paragraphProperties6.Append(paragraphMarkRunProperties6);

                Run run7 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties7 = new RunProperties();
                RunFonts runFonts13 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize6 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };
                Languages languages4 = new Languages() { EastAsia = "zh-HK" };

                runProperties7.Append(runFonts13);
                runProperties7.Append(fontSize6);
                runProperties7.Append(fontSizeComplexScript6);
                runProperties7.Append(languages4);
                Text text7 = new Text();
                text7.Text = "單位主管您好";

                run7.Append(runProperties7);
                run7.Append(text7);

                Run run8 = new Run() { RsidRunProperties = "00691BB8", RsidRunAddition = "00CA3D7A" };

                RunProperties runProperties8 = new RunProperties();
                RunFonts runFonts14 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize7 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

                runProperties8.Append(runFonts14);
                runProperties8.Append(fontSize7);
                runProperties8.Append(fontSizeComplexScript7);
                Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text8.Text = ", ";

                run8.Append(runProperties8);
                run8.Append(text8);

                paragraph6.Append(paragraphProperties6);
                paragraph6.Append(run7);
                paragraph6.Append(run8);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00CA3D7A", RsidParagraphProperties = "0050595B", RsidRunAdditionDefault = "00CA3D7A", ParagraphId = "67BEC247", TextId = "77777777" };

                ParagraphProperties paragraphProperties7 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
                RunFonts runFonts15 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize8 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties7.Append(runFonts15);
                paragraphMarkRunProperties7.Append(fontSize8);
                paragraphMarkRunProperties7.Append(fontSizeComplexScript8);

                paragraphProperties7.Append(spacingBetweenLines3);
                paragraphProperties7.Append(paragraphMarkRunProperties7);

                paragraph7.Append(paragraphProperties7);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00FA79CC", RsidParagraphProperties = "0050595B", RsidRunAdditionDefault = "00C747E9", ParagraphId = "1C30160E", TextId = "437B26FA" };

                ParagraphProperties paragraphProperties8 = new ParagraphProperties();
                SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation1 = new Indentation() { End = "-24", EndCharacters = -10 };
                Justification justification5 = new Justification() { Val = JustificationValues.Both };

                ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
                RunFonts runFonts16 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize9 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties8.Append(runFonts16);
                paragraphMarkRunProperties8.Append(fontSize9);
                paragraphMarkRunProperties8.Append(fontSizeComplexScript9);

                paragraphProperties8.Append(spacingBetweenLines4);
                paragraphProperties8.Append(indentation1);
                paragraphProperties8.Append(justification5);
                paragraphProperties8.Append(paragraphMarkRunProperties8);

                Run run9 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties9 = new RunProperties();
                RunFonts runFonts17 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize10 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "32" };
                Languages languages5 = new Languages() { EastAsia = "zh-HK" };

                runProperties9.Append(runFonts17);
                runProperties9.Append(fontSize10);
                runProperties9.Append(fontSizeComplexScript10);
                runProperties9.Append(languages5);
                Text text9 = new Text();
                text9.Text = "茲";

                run9.Append(runProperties9);
                run9.Append(text9);

                Run run10 = new Run() { RsidRunProperties = "4CC70866", RsidRunAddition = "00A02E9E" };

                RunProperties runProperties10 = new RunProperties();
                RunFonts runFonts18 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize11 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "32" };
                Languages languages6 = new Languages() { EastAsia = "zh-HK" };

                runProperties10.Append(runFonts18);
                runProperties10.Append(fontSize11);
                runProperties10.Append(fontSizeComplexScript11);
                runProperties10.Append(languages6);
                Text text10 = new Text();
                text10.Text = "依據「金融控股公司及銀行業內部控制及稽核制度實施辦法」";

                run10.Append(runProperties10);
                run10.Append(text10);

                Run run11 = new Run() { RsidRunProperties = "4CC70866", RsidRunAddition = "00FA79CC" };

                RunProperties runProperties11 = new RunProperties();
                RunFonts runFonts19 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize12 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "32" };

                runProperties11.Append(runFonts19);
                runProperties11.Append(fontSize12);
                runProperties11.Append(fontSizeComplexScript12);
                Text text11 = new Text();
                text11.Text = "，前往貴單位";

                run11.Append(runProperties11);
                run11.Append(text11);

                Run run12 = new Run() { RsidRunProperties = "4CC70866", RsidRunAddition = "00FA79CC" };

                RunProperties runProperties12 = new RunProperties();
                RunFonts runFonts20 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize13 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "32" };
                Languages languages7 = new Languages() { EastAsia = "zh-HK" };

                runProperties12.Append(runFonts20);
                runProperties12.Append(fontSize13);
                runProperties12.Append(fontSizeComplexScript13);
                runProperties12.Append(languages7);
                Text text12 = new Text();
                text12.Text = "辦理";

                run12.Append(runProperties12);
                run12.Append(text12);

                Run run13 = new Run() { RsidRunProperties = "4CC70866", RsidRunAddition = "00FA79CC" };

                RunProperties runProperties13 = new RunProperties();
                RunFonts runFonts21 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize14 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "32" };

                runProperties13.Append(runFonts21);
                runProperties13.Append(fontSize14);
                runProperties13.Append(fontSizeComplexScript14);
                Text text13 = new Text();
                text13.Text = "檢查，即請";

                run13.Append(runProperties13);
                run13.Append(text13);
                ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run14 = new Run() { RsidRunProperties = "4CC70866", RsidRunAddition = "00FA79CC" };

                RunProperties runProperties14 = new RunProperties();
                RunFonts runFonts22 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize15 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "32" };

                runProperties14.Append(runFonts22);
                runProperties14.Append(fontSize15);
                runProperties14.Append(fontSizeComplexScript15);
                Text text14 = new Text();
                text14.Text = "查照並惠予";

                run14.Append(runProperties14);
                run14.Append(text14);
                ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run15 = new Run() { RsidRunProperties = "4CC70866", RsidRunAddition = "00FA79CC" };

                RunProperties runProperties15 = new RunProperties();
                RunFonts runFonts23 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize16 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "32" };

                runProperties15.Append(runFonts23);
                runProperties15.Append(fontSize16);
                runProperties15.Append(fontSizeComplexScript16);
                Text text15 = new Text();
                text15.Text = "協助為荷";

                run15.Append(runProperties15);
                run15.Append(text15);

                Run run16 = new Run() { RsidRunAddition = "0086389F" };

                RunProperties runProperties16 = new RunProperties();
                RunFonts runFonts24 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize17 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "32" };

                runProperties16.Append(runFonts24);
                runProperties16.Append(fontSize17);
                runProperties16.Append(fontSizeComplexScript17);
                Text text16 = new Text();
                text16.Text = "。";

                run16.Append(runProperties16);
                run16.Append(text16);

                paragraph8.Append(paragraphProperties8);
                paragraph8.Append(run9);
                paragraph8.Append(run10);
                paragraph8.Append(run11);
                paragraph8.Append(run12);
                paragraph8.Append(run13);
                paragraph8.Append(proofError1);
                paragraph8.Append(run14);
                paragraph8.Append(proofError2);
                paragraph8.Append(run15);
                paragraph8.Append(run16);

                Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "1FB659BE", RsidParagraphProperties = "0086389F", RsidRunAdditionDefault = "1FB659BE", ParagraphId = "1F244CF4", TextId = "1F3ED55B" };

                ParagraphProperties paragraphProperties9 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "ac" };

                NumberingProperties numberingProperties1 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
                NumberingId numberingId1 = new NumberingId() { Val = 1 };

                numberingProperties1.Append(numberingLevelReference1);
                numberingProperties1.Append(numberingId1);
                SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "360", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation2 = new Indentation() { Start = "964", StartCharacters = 0, End = "-58", EndCharacters = -24, Hanging = "482" };

                ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
                RunFonts runFonts25 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                FontSize fontSize18 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "32" };

                paragraphMarkRunProperties9.Append(runFonts25);
                paragraphMarkRunProperties9.Append(fontSize18);
                paragraphMarkRunProperties9.Append(fontSizeComplexScript18);

                paragraphProperties9.Append(paragraphStyleId1);
                paragraphProperties9.Append(numberingProperties1);
                paragraphProperties9.Append(spacingBetweenLines5);
                paragraphProperties9.Append(indentation2);
                paragraphProperties9.Append(paragraphMarkRunProperties9);
                ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run17 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties17 = new RunProperties();
                RunFonts runFonts26 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize19 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "32" };
                Languages languages8 = new Languages() { EastAsia = "zh-HK" };

                runProperties17.Append(runFonts26);
                runProperties17.Append(fontSize19);
                runProperties17.Append(fontSizeComplexScript19);
                runProperties17.Append(languages8);
                Text text17 = new Text();
                text17.Text = "查程編號";

                run17.Append(runProperties17);
                run17.Append(text17);
                ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run18 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties18 = new RunProperties();
                RunFonts runFonts27 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                FontSize fontSize20 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "32" };

                runProperties18.Append(runFonts27);
                runProperties18.Append(fontSize20);
                runProperties18.Append(fontSizeComplexScript20);
                Text text18 = new Text();
                text18.Text = "：";

                run18.Append(runProperties18);
                run18.Append(text18);

                Run run19 = new Run() { RsidRunProperties = "0086389F", RsidRunAddition = "0086389F" };

                RunProperties runProperties19 = new RunProperties();
                RunFonts runFonts28 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體" };
                Color color5 = new Color() { Val = "0000FF" };
                FontSize fontSize21 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "32" };

                runProperties19.Append(runFonts28);
                runProperties19.Append(color5);
                runProperties19.Append(fontSize21);
                runProperties19.Append(fontSizeComplexScript21);
                Text text19 = new Text();
                text19.Text = rows["CompanyName"].ToString();

                run19.Append(runProperties19);
                run19.Append(text19);

                Run run20 = new Run() { RsidRunProperties = "0086389F", RsidRunAddition = "0086389F" };

                RunProperties runProperties20 = new RunProperties();
                RunFonts runFonts29 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體" };
                Color color6 = new Color() { Val = "0000FF" };
                FontSize fontSize22 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "32" };

                runProperties20.Append(runFonts29);
                runProperties20.Append(color6);
                runProperties20.Append(fontSize22);
                runProperties20.Append(fontSizeComplexScript22);
                Text text20 = new Text();
                text20.Text = "_";

                run20.Append(runProperties20);
                run20.Append(text20);

                Run run21 = new Run() { RsidRunProperties = "4CC70866" };

                RunProperties runProperties21 = new RunProperties();
                RunFonts runFonts30 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color7 = new Color() { Val = "0000FF" };
                FontSize fontSize23 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "32" };

                runProperties21.Append(runFonts30);
                runProperties21.Append(color7);
                runProperties21.Append(fontSize23);
                runProperties21.Append(fontSizeComplexScript23);
                Text text21 = new Text();
                text21.Text = rows["audit_no"].ToString();

                run21.Append(runProperties21);
                run21.Append(text21);

                paragraph9.Append(paragraphProperties9);
                paragraph9.Append(proofError3);
                paragraph9.Append(run17);
                paragraph9.Append(proofError4);
                paragraph9.Append(run18);
                paragraph9.Append(run19);
                paragraph9.Append(run20);
                paragraph9.Append(run21);

                Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00104B80", RsidParagraphProperties = "0086389F", RsidRunAdditionDefault = "00104B80", ParagraphId = "7BCADF04", TextId = "2B38C043" };

                ParagraphProperties paragraphProperties10 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "ac" };

                NumberingProperties numberingProperties2 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference2 = new NumberingLevelReference() { Val = 0 };
                NumberingId numberingId2 = new NumberingId() { Val = 1 };

                numberingProperties2.Append(numberingLevelReference2);
                numberingProperties2.Append(numberingId2);
                AdjustRightIndent adjustRightIndent1 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };
                SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation3 = new Indentation() { Start = "964", StartCharacters = 0, End = "557", EndCharacters = 232, Hanging = "482" };

                ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
                RunFonts runFonts31 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern1 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize24 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties10.Append(runFonts31);
                paragraphMarkRunProperties10.Append(kern1);
                paragraphMarkRunProperties10.Append(fontSize24);
                paragraphMarkRunProperties10.Append(fontSizeComplexScript24);

                paragraphProperties10.Append(paragraphStyleId2);
                paragraphProperties10.Append(numberingProperties2);
                paragraphProperties10.Append(adjustRightIndent1);
                paragraphProperties10.Append(snapToGrid1);
                paragraphProperties10.Append(spacingBetweenLines6);
                paragraphProperties10.Append(indentation3);
                paragraphProperties10.Append(paragraphMarkRunProperties10);
                ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run22 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties22 = new RunProperties();
                RunFonts runFonts32 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize25 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "32" };
                Languages languages9 = new Languages() { EastAsia = "zh-HK" };

                runProperties22.Append(runFonts32);
                runProperties22.Append(fontSize25);
                runProperties22.Append(fontSizeComplexScript25);
                runProperties22.Append(languages9);
                Text text22 = new Text();
                text22.Text = "查程名稱";

                run22.Append(runProperties22);
                run22.Append(text22);
                ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run23 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties23 = new RunProperties();
                RunFonts runFonts33 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize26 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "32" };

                runProperties23.Append(runFonts33);
                runProperties23.Append(fontSize26);
                runProperties23.Append(fontSizeComplexScript26);
                Text text23 = new Text();
                text23.Text = "：";

                run23.Append(runProperties23);
                run23.Append(text23);

                Run run24 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties24 = new RunProperties();
                RunFonts runFonts34 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color8 = new Color() { Val = "0000FF" };
                Kern kern2 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize27 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "32" };

                runProperties24.Append(runFonts34);
                runProperties24.Append(color8);
                runProperties24.Append(kern2);
                runProperties24.Append(fontSize27);
                runProperties24.Append(fontSizeComplexScript27);
                Text text24 = new Text();
                text24.Text = rows["planname"].ToString();

                run24.Append(runProperties24);
                run24.Append(text24);

                Run run25 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties25 = new RunProperties();
                RunFonts runFonts35 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color9 = new Color() { Val = "0000FF" };
                Kern kern3 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize28 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "32" };

                runProperties25.Append(runFonts35);
                runProperties25.Append(color9);
                runProperties25.Append(kern3);
                runProperties25.Append(fontSize28);
                runProperties25.Append(fontSizeComplexScript28);
                Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text25.Text = " ";

                run25.Append(runProperties25);
                run25.Append(text25);

                Run run26 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties26 = new RunProperties();
                RunFonts runFonts36 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern4 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize29 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "32" };

                runProperties26.Append(runFonts36);
                runProperties26.Append(kern4);
                runProperties26.Append(fontSize29);
                runProperties26.Append(fontSizeComplexScript29);
                Text text26 = new Text();
                text26.Text = "(";

                run26.Append(runProperties26);
                run26.Append(text26);

                Run run27 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties27 = new RunProperties();
                RunFonts runFonts37 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color10 = new Color() { Val = "0000FF" };
                Kern kern5 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize30 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "32" };

                runProperties27.Append(runFonts37);
                runProperties27.Append(color10);
                runProperties27.Append(kern5);
                runProperties27.Append(fontSize30);
                runProperties27.Append(fontSizeComplexScript30);
                Text text27 = new Text();
                text27.Text = "";

                run27.Append(runProperties27);
                run27.Append(text27);

                Run run28 = new Run() { RsidRunAddition = "00080C65" };

                RunProperties runProperties28 = new RunProperties();
                RunFonts runFonts38 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color11 = new Color() { Val = "0000FF" };
                Kern kern6 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize31 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "32" };
                Languages languages10 = new Languages() { EastAsia = "zh-HK" };

                runProperties28.Append(runFonts38);
                runProperties28.Append(color11);
                runProperties28.Append(kern6);
                runProperties28.Append(fontSize31);
                runProperties28.Append(fontSizeComplexScript31);
                runProperties28.Append(languages10);
                Text text28 = new Text();
                text28.Text = rows["plantype"].ToString();

                run28.Append(runProperties28);
                run28.Append(text28);

                Run run29 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties29 = new RunProperties();
                RunFonts runFonts39 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern7 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize32 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "32" };

                runProperties29.Append(runFonts39);
                runProperties29.Append(kern7);
                runProperties29.Append(fontSize32);
                runProperties29.Append(fontSizeComplexScript32);
                Text text29 = new Text();
                text29.Text = ")";

                run29.Append(runProperties29);
                run29.Append(text29);

                paragraph10.Append(paragraphProperties10);
                paragraph10.Append(proofError5);
                paragraph10.Append(run22);
                paragraph10.Append(proofError6);
                paragraph10.Append(run23);
                paragraph10.Append(run24);
                paragraph10.Append(run25);
                paragraph10.Append(run26);
                paragraph10.Append(run27);
                paragraph10.Append(run28);
                paragraph10.Append(run29);

                Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00104B80", RsidParagraphProperties = "00D66A62", RsidRunAdditionDefault = "00104B80", ParagraphId = "30DA84AE", TextId = "1736F17B" };

                ParagraphProperties paragraphProperties11 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "ac" };

                NumberingProperties numberingProperties3 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference3 = new NumberingLevelReference() { Val = 0 };
                NumberingId numberingId3 = new NumberingId() { Val = 1 };

                numberingProperties3.Append(numberingLevelReference3);
                numberingProperties3.Append(numberingId3);
                AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };
                SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation4 = new Indentation() { StartCharacters = 0, End = "557", EndCharacters = 232 };

                ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
                RunFonts runFonts40 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                FontSize fontSize33 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties11.Append(runFonts40);
                paragraphMarkRunProperties11.Append(fontSize33);
                paragraphMarkRunProperties11.Append(fontSizeComplexScript33);

                paragraphProperties11.Append(paragraphStyleId3);
                paragraphProperties11.Append(numberingProperties3);
                paragraphProperties11.Append(adjustRightIndent2);
                paragraphProperties11.Append(snapToGrid2);
                paragraphProperties11.Append(spacingBetweenLines7);
                paragraphProperties11.Append(indentation4);
                paragraphProperties11.Append(paragraphMarkRunProperties11);

                Run run30 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties30 = new RunProperties();
                RunFonts runFonts41 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern8 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize34 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "32" };
                Languages languages11 = new Languages() { EastAsia = "zh-HK" };

                runProperties30.Append(runFonts41);
                runProperties30.Append(kern8);
                runProperties30.Append(fontSize34);
                runProperties30.Append(fontSizeComplexScript34);
                runProperties30.Append(languages11);
                Text text30 = new Text();
                text30.Text = "查核期間";

                run30.Append(runProperties30);
                run30.Append(text30);

                Run run31 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties31 = new RunProperties();
                RunFonts runFonts42 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern9 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize35 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "32" };

                runProperties31.Append(runFonts42);
                runProperties31.Append(kern9);
                runProperties31.Append(fontSize35);
                runProperties31.Append(fontSizeComplexScript35);
                Text text31 = new Text();
                text31.Text = "：";

                run31.Append(runProperties31);
                run31.Append(text31);

                Run run32 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties32 = new RunProperties();
                RunFonts runFonts43 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color12 = new Color() { Val = "0000FF" };
                Kern kern10 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize36 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "32" };

                runProperties32.Append(runFonts43);
                runProperties32.Append(color12);
                runProperties32.Append(kern10);
                runProperties32.Append(fontSize36);
                runProperties32.Append(fontSizeComplexScript36);
                Text text32 = new Text();
                text32.Text = Convert.ToDateTime(rows["startdate"]).ToString("yyyy-MM-dd");

                run32.Append(runProperties32);
                run32.Append(text32);

                Run run33 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties33 = new RunProperties();
                RunFonts runFonts44 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color13 = new Color() { Val = "0000FF" };
                Kern kern11 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize37 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "32" };

                runProperties33.Append(runFonts44);
                runProperties33.Append(color13);
                runProperties33.Append(kern11);
                runProperties33.Append(fontSize37);
                runProperties33.Append(fontSizeComplexScript37);
                Text text33 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text33.Text = " ~ ";

                run33.Append(runProperties33);
                run33.Append(text33);

                Run run34 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties34 = new RunProperties();
                RunFonts runFonts45 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color14 = new Color() { Val = "0000FF" };
                Kern kern12 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize38 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "32" };

                runProperties34.Append(runFonts45);
                runProperties34.Append(color14);
                runProperties34.Append(kern12);
                runProperties34.Append(fontSize38);
                runProperties34.Append(fontSizeComplexScript38);
                Text text34 = new Text();
                text34.Text = Convert.ToDateTime(rows["enddate"]).ToString("yyyy-MM-dd");

                run34.Append(runProperties34);
                run34.Append(text34);

                Run run35 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties35 = new RunProperties();
                RunFonts runFonts46 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color15 = new Color() { Val = "000000" };
                Kern kern13 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize39 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "32" };

                runProperties35.Append(runFonts46);
                runProperties35.Append(color15);
                runProperties35.Append(kern13);
                runProperties35.Append(fontSize39);
                runProperties35.Append(fontSizeComplexScript39);
                FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

                run35.Append(runProperties35);
                run35.Append(fieldChar1);

                Run run36 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties36 = new RunProperties();
                RunFonts runFonts47 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color16 = new Color() { Val = "000000" };
                Kern kern14 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize40 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "32" };

                runProperties36.Append(runFonts47);
                runProperties36.Append(color16);
                runProperties36.Append(kern14);
                runProperties36.Append(fontSize40);
                runProperties36.Append(fontSizeComplexScript40);
                FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
                fieldCode1.Text = " DOCVARIABLE  qORG  \\* MERGEFORMAT ";

                run36.Append(runProperties36);
                run36.Append(fieldCode1);

                Run run37 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties37 = new RunProperties();
                RunFonts runFonts48 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color17 = new Color() { Val = "000000" };
                Kern kern15 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize41 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "32" };

                runProperties37.Append(runFonts48);
                runProperties37.Append(color17);
                runProperties37.Append(kern15);
                runProperties37.Append(fontSize41);
                runProperties37.Append(fontSizeComplexScript41);
                FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.End };

                run37.Append(runProperties37);
                run37.Append(fieldChar2);

                paragraph11.Append(paragraphProperties11);
                paragraph11.Append(run30);
                paragraph11.Append(run31);
                paragraph11.Append(run32);
                paragraph11.Append(run33);
                paragraph11.Append(run34);
                paragraph11.Append(run35);
                paragraph11.Append(run36);
                paragraph11.Append(run37);

                Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00104B80", RsidParagraphProperties = "00D66A62", RsidRunAdditionDefault = "00104B80", ParagraphId = "4A20D429", TextId = "77777777" };

                ParagraphProperties paragraphProperties12 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "ac" };

                NumberingProperties numberingProperties4 = new NumberingProperties();
                NumberingLevelReference numberingLevelReference4 = new NumberingLevelReference() { Val = 0 };
                NumberingId numberingId4 = new NumberingId() { Val = 1 };

                numberingProperties4.Append(numberingLevelReference4);
                numberingProperties4.Append(numberingId4);
                AdjustRightIndent adjustRightIndent3 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid3 = new SnapToGrid() { Val = false };
                SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation5 = new Indentation() { StartCharacters = 0, End = "557", EndCharacters = 232 };

                ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
                RunFonts runFonts49 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern16 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize42 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "28" };
                Languages languages12 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties12.Append(runFonts49);
                paragraphMarkRunProperties12.Append(kern16);
                paragraphMarkRunProperties12.Append(fontSize42);
                paragraphMarkRunProperties12.Append(fontSizeComplexScript42);
                paragraphMarkRunProperties12.Append(languages12);

                paragraphProperties12.Append(paragraphStyleId4);
                paragraphProperties12.Append(numberingProperties4);
                paragraphProperties12.Append(adjustRightIndent3);
                paragraphProperties12.Append(snapToGrid3);
                paragraphProperties12.Append(spacingBetweenLines8);
                paragraphProperties12.Append(indentation5);
                paragraphProperties12.Append(paragraphMarkRunProperties12);

                Run run38 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties38 = new RunProperties();
                RunFonts runFonts50 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern17 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize43 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "32" };
                Languages languages13 = new Languages() { EastAsia = "zh-HK" };

                runProperties38.Append(runFonts50);
                runProperties38.Append(kern17);
                runProperties38.Append(fontSize43);
                runProperties38.Append(fontSizeComplexScript43);
                runProperties38.Append(languages13);
                Text text35 = new Text();
                text35.Text = "查核成員";

                run38.Append(runProperties38);
                run38.Append(text35);

                Run run39 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties39 = new RunProperties();
                RunFonts runFonts51 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern18 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize44 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "32" };

                runProperties39.Append(runFonts51);
                runProperties39.Append(kern18);
                runProperties39.Append(fontSize44);
                runProperties39.Append(fontSizeComplexScript44);
                Text text36 = new Text();
                text36.Text = "：";

                run39.Append(runProperties39);
                run39.Append(text36);

                paragraph12.Append(paragraphProperties12);
                paragraph12.Append(run38);
                paragraph12.Append(run39);

                Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00CA3D7A", RsidParagraphProperties = "00D66A62", RsidRunAdditionDefault = "00FA79CC", ParagraphId = "687817A3", TextId = "2EF438C1" };

                ParagraphProperties paragraphProperties13 = new ParagraphProperties();
                AdjustRightIndent adjustRightIndent4 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid4 = new SnapToGrid() { Val = false };
                SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation6 = new Indentation() { Start = "960", End = "557", EndCharacters = 232 };

                ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
                RunFonts runFonts52 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern19 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize45 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "28" };
                Languages languages14 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties13.Append(runFonts52);
                paragraphMarkRunProperties13.Append(kern19);
                paragraphMarkRunProperties13.Append(fontSize45);
                paragraphMarkRunProperties13.Append(fontSizeComplexScript45);
                paragraphMarkRunProperties13.Append(languages14);

                paragraphProperties13.Append(adjustRightIndent4);
                paragraphProperties13.Append(snapToGrid4);
                paragraphProperties13.Append(spacingBetweenLines9);
                paragraphProperties13.Append(indentation6);
                paragraphProperties13.Append(paragraphMarkRunProperties13);

                Run run40 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties40 = new RunProperties();
                RunFonts runFonts53 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize46 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "28" };

                runProperties40.Append(runFonts53);
                runProperties40.Append(fontSize46);
                runProperties40.Append(fontSizeComplexScript46);
                Text text37 = new Text();
                text37.Text = "領隊";

                run40.Append(runProperties40);
                run40.Append(text37);

                Run run41 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties41 = new RunProperties();
                RunFonts runFonts54 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern20 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize47 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "28" };
                Languages languages15 = new Languages() { EastAsia = "zh-HK" };

                runProperties41.Append(runFonts54);
                runProperties41.Append(kern20);
                runProperties41.Append(fontSize47);
                runProperties41.Append(fontSizeComplexScript47);
                runProperties41.Append(languages15);
                Text text38 = new Text();
                text38.Text = "稽核";

                run41.Append(runProperties41);
                run41.Append(text38);

                Run run42 = new Run() { RsidRunProperties = "00691BB8", RsidRunAddition = "00104B80" };

                RunProperties runProperties42 = new RunProperties();
                RunFonts runFonts55 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern21 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize48 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "28" };

                runProperties42.Append(runFonts55);
                runProperties42.Append(kern21);
                runProperties42.Append(fontSize48);
                runProperties42.Append(fontSizeComplexScript48);
                Text text39 = new Text();
                text39.Text = "：";

                run42.Append(runProperties42);
                run42.Append(text39);

                Run run43 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties43 = new RunProperties();
                RunFonts runFonts56 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Color color18 = new Color() { Val = "0000FF" };
                Kern kern22 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize49 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "28" };
                Languages languages16 = new Languages() { EastAsia = "zh-HK" };

                runProperties43.Append(runFonts56);
                runProperties43.Append(color18);
                runProperties43.Append(kern22);
                runProperties43.Append(fontSize49);
                runProperties43.Append(fontSizeComplexScript49);
                runProperties43.Append(languages16);
                Text text40 = new Text();
                text40.Text = rows["leader"].ToString();

                run43.Append(runProperties43);
                run43.Append(text40);

                paragraph13.Append(paragraphProperties13);
                paragraph13.Append(run40);
                paragraph13.Append(run41);
                paragraph13.Append(run42);
                paragraph13.Append(run43);

                Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00CA3D7A", RsidParagraphProperties = "00292721", RsidRunAdditionDefault = "00FA79CC", ParagraphId = "622FB3E3", TextId = "1FD669C3" };

                ParagraphProperties paragraphProperties14 = new ParagraphProperties();
                AdjustRightIndent adjustRightIndent5 = new AdjustRightIndent() { Val = false };
                SnapToGrid snapToGrid5 = new SnapToGrid() { Val = false };
                SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "360", LineRule = LineSpacingRuleValues.Auto };
                Indentation indentation7 = new Indentation() { Start = "1440", StartCharacters = 600, End = "557", EndCharacters = 232, FirstLine = "160", FirstLineChars = 50 };

                ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
                RunFonts runFonts57 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern23 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize50 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "28" };
                Languages languages17 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties14.Append(runFonts57);
                paragraphMarkRunProperties14.Append(kern23);
                paragraphMarkRunProperties14.Append(fontSize50);
                paragraphMarkRunProperties14.Append(fontSizeComplexScript50);
                paragraphMarkRunProperties14.Append(languages17);

                paragraphProperties14.Append(adjustRightIndent5);
                paragraphProperties14.Append(snapToGrid5);
                paragraphProperties14.Append(spacingBetweenLines10);
                paragraphProperties14.Append(indentation7);
                paragraphProperties14.Append(paragraphMarkRunProperties14);

                Run run44 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties44 = new RunProperties();
                RunFonts runFonts58 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern24 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize51 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "28" };
                Languages languages18 = new Languages() { EastAsia = "zh-HK" };

                runProperties44.Append(runFonts58);
                runProperties44.Append(kern24);
                runProperties44.Append(fontSize51);
                runProperties44.Append(fontSizeComplexScript51);
                runProperties44.Append(languages18);
                Text text41 = new Text();
                text41.Text = "稽核";

                run44.Append(runProperties44);
                run44.Append(text41);

                Run run45 = new Run() { RsidRunAddition = "007607FB" };

                RunProperties runProperties45 = new RunProperties();
                RunFonts runFonts59 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                Kern kern25 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize52 = new FontSize() { Val = "32" };
                FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "28" };

                runProperties45.Append(runFonts59);
                runProperties45.Append(kern25);
                runProperties45.Append(fontSize52);
                runProperties45.Append(fontSizeComplexScript52);
                Text text42 = new Text();
                text42.Text = "：";

                run45.Append(runProperties45);
                run45.Append(text42);

                paragraph14.Append(paragraphProperties14);
                paragraph14.Append(run44);
                paragraph14.Append(run45);
                string[] MemberArr = rows["Member"].ToString().Split(',');
                foreach (string Member in MemberArr)
                {
                    Run run46 = new Run() { RsidRunProperties = "00691BB8" };

                    RunProperties runProperties46 = new RunProperties();
                    RunFonts runFonts60 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "標楷體" };
                    Color color19 = new Color() { Val = "0000FF" };
                    Kern kern26 = new Kern() { Val = (UInt32Value)0U };
                    FontSize fontSize53 = new FontSize() { Val = "32" };
                    FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "28" };
                    Languages languages19 = new Languages() { EastAsia = "zh-HK" };

                    runProperties46.Append(runFonts60);
                    runProperties46.Append(color19);
                    runProperties46.Append(kern26);
                    runProperties46.Append(fontSize53);
                    runProperties46.Append(fontSizeComplexScript53);
                    runProperties46.Append(languages19);
                    Text text43 = new Text();
                    text43.Text = Member;

                    Break lineBreak = new Break();
                    run46.Append(runProperties46);
                    run46.Append(text43);

                    paragraph14.Append(run46);
                    paragraph14.Append(lineBreak);
                }



                Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "00691BB8", RsidRunAdditionDefault = "00691BB8", ParagraphId = "35B377F3", TextId = "77777777" };

                ParagraphProperties paragraphProperties17 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
                RunFonts runFonts65 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize57 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "28" };
                Languages languages22 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties17.Append(runFonts65);
                paragraphMarkRunProperties17.Append(fontSize57);
                paragraphMarkRunProperties17.Append(fontSizeComplexScript58);
                paragraphMarkRunProperties17.Append(languages22);

                paragraphProperties17.Append(paragraphMarkRunProperties17);

                paragraph17.Append(paragraphProperties17);

                Paragraph paragraph18 = new Paragraph() { RsidParagraphAddition = "00691BB8", RsidRunAdditionDefault = "00691BB8", ParagraphId = "410D77B3", TextId = "77777777" };

                ParagraphProperties paragraphProperties18 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
                RunFonts runFonts66 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize58 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "28" };
                Languages languages23 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties18.Append(runFonts66);
                paragraphMarkRunProperties18.Append(fontSize58);
                paragraphMarkRunProperties18.Append(fontSizeComplexScript59);
                paragraphMarkRunProperties18.Append(languages23);

                paragraphProperties18.Append(paragraphMarkRunProperties18);

                paragraph18.Append(paragraphProperties18);

                Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "00691BB8", RsidRunAdditionDefault = "00691BB8", ParagraphId = "286D8E41", TextId = "77777777" };

                ParagraphProperties paragraphProperties19 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
                RunFonts runFonts67 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize59 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "28" };
                Languages languages24 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties19.Append(runFonts67);
                paragraphMarkRunProperties19.Append(fontSize59);
                paragraphMarkRunProperties19.Append(fontSizeComplexScript60);
                paragraphMarkRunProperties19.Append(languages24);

                paragraphProperties19.Append(paragraphMarkRunProperties19);

                paragraph19.Append(paragraphProperties19);

                Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "00691BB8", RsidRunAdditionDefault = "00691BB8", ParagraphId = "0DF7231E", TextId = "77777777" };

                ParagraphProperties paragraphProperties20 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
                RunFonts runFonts68 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                FontSize fontSize60 = new FontSize() { Val = "28" };
                FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "28" };
                Languages languages25 = new Languages() { EastAsia = "zh-HK" };

                paragraphMarkRunProperties20.Append(runFonts68);
                paragraphMarkRunProperties20.Append(fontSize60);
                paragraphMarkRunProperties20.Append(fontSizeComplexScript61);
                paragraphMarkRunProperties20.Append(languages25);

                paragraphProperties20.Append(paragraphMarkRunProperties20);

                paragraph20.Append(paragraphProperties20);

                Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00691BB8", RsidParagraphAddition = "00E76DB3", RsidRunAdditionDefault = "00691BB8", ParagraphId = "5D0ADEA8", TextId = "0D296A86" };

                ParagraphProperties paragraphProperties21 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
                RunFonts runFonts69 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color23 = new Color() { Val = "0000FF" };
                FontSize fontSize61 = new FontSize() { Val = "44" };
                FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties21.Append(runFonts69);
                paragraphMarkRunProperties21.Append(color23);
                paragraphMarkRunProperties21.Append(fontSize61);
                paragraphMarkRunProperties21.Append(fontSizeComplexScript62);

                paragraphProperties21.Append(paragraphMarkRunProperties21);

                Run run49 = new Run() { RsidRunProperties = "00691BB8" };

                RunProperties runProperties49 = new RunProperties();
                RunFonts runFonts70 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color24 = new Color() { Val = "0000FF" };
                FontSize fontSize62 = new FontSize() { Val = "44" };
                FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "28" };
                Languages languages26 = new Languages() { EastAsia = "zh-HK" };

                runProperties49.Append(runFonts70);
                runProperties49.Append(color24);
                runProperties49.Append(fontSize62);
                runProperties49.Append(fontSizeComplexScript63);
                runProperties49.Append(languages26);
                Text text46 = new Text();
                text46.Text = rows["Belong"].ToString();

                run49.Append(runProperties49);
                run49.Append(text46);

                paragraph21.Append(paragraphProperties21);
                paragraph21.Append(run49);

                SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "00691BB8", RsidR = "00E76DB3", RsidSect = "00FA79CC" };
                HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId11" };
                PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
                PageMargin pageMargin1 = new PageMargin() { Top = 1440, Right = (UInt32Value)1800U, Bottom = 1440, Left = (UInt32Value)1800U, Header = (UInt32Value)851U, Footer = (UInt32Value)992U, Gutter = (UInt32Value)0U };
                Columns columns1 = new Columns() { Space = "425" };
                DocGrid docGrid1 = new DocGrid() { Type = DocGridValues.Lines, LinePitch = 360 };

                sectionProperties1.Append(headerReference1);
                sectionProperties1.Append(pageSize1);
                sectionProperties1.Append(pageMargin1);
                sectionProperties1.Append(columns1);
                sectionProperties1.Append(docGrid1);

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
                //body1.Append(paragraph16);
                //body1.Append(paragraph17);
                //body1.Append(paragraph18);
                body1.Append(paragraph19);
                body1.Append(paragraph20);
                body1.Append(paragraph21);
                body1.Append(sectionProperties1);
                document1.Append(body1);
            }

            mainDocumentPart1.Document = document1;
            }

            // Generates content of webSettingsPart1.
            private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
            {
                WebSettings webSettings1 = new WebSettings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
                webSettings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                webSettings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                webSettings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                webSettings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                webSettings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                webSettings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                webSettings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                webSettings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

                Divs divs1 = new Divs();

                Div div1 = new Div() { Id = "2000384593" };
                BodyDiv bodyDiv1 = new BodyDiv() { Val = true };
                LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv() { Val = "0" };
                RightMarginDiv rightMarginDiv1 = new RightMarginDiv() { Val = "0" };
                TopMarginDiv topMarginDiv1 = new TopMarginDiv() { Val = "0" };
                BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv() { Val = "0" };

                DivBorder divBorder1 = new DivBorder();
                TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
                RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

                divBorder1.Append(topBorder1);
                divBorder1.Append(leftBorder1);
                divBorder1.Append(bottomBorder1);
                divBorder1.Append(rightBorder1);

                div1.Append(bodyDiv1);
                div1.Append(leftMarginDiv1);
                div1.Append(rightMarginDiv1);
                div1.Append(topMarginDiv1);
                div1.Append(bottomMarginDiv1);
                div1.Append(divBorder1);

                divs1.Append(div1);
                OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();
                AllowPNG allowPNG1 = new AllowPNG();

                webSettings1.Append(divs1);
                webSettings1.Append(optimizeForBrowser1);
                webSettings1.Append(allowPNG1);

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
                A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

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
                majorFont1.Append(supplementalFont30);

                A.MinorFont minorFont1 = new A.MinorFont();
                A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
                A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
                A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
                A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
                A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
                A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
                A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
                A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
                A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
                A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
                A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
                A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
                A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
                A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
                A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
                A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
                A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
                A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
                A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
                A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
                A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
                A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
                A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
                A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
                A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
                A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
                A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
                A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
                A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
                A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
                A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
                A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
                A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

                minorFont1.Append(latinFont2);
                minorFont1.Append(eastAsianFont2);
                minorFont1.Append(complexScriptFont2);
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
                minorFont1.Append(supplementalFont59);
                minorFont1.Append(supplementalFont60);

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

            // Generates content of customXmlPart1.
            private void GenerateCustomXmlPart1Content(CustomXmlPart customXmlPart1)
            {
                System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><p:properties xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"><documentManagement/></p:properties>");
                writer.Flush();
                writer.Close();
            }

            // Generates content of customXmlPropertiesPart1.
            private void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
            {
                Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem() { ItemId = "{8F7B3920-A7CD-4BCE-8EC3-E7D228EB3CE6}" };
                dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

                Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
                Ds.SchemaReference schemaReference1 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties" };
                Ds.SchemaReference schemaReference2 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls" };

                schemaReferences1.Append(schemaReference1);
                schemaReferences1.Append(schemaReference2);

                dataStoreItem1.Append(schemaReferences1);

                customXmlPropertiesPart1.DataStoreItem = dataStoreItem1;
            }

            // Generates content of documentSettingsPart1.
            private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
            {
                Settings settings1 = new Settings() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
                settings1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                settings1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                settings1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                settings1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                settings1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                settings1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                settings1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                settings1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
                Zoom zoom1 = new Zoom() { Percent = "130" };
                BordersDoNotSurroundHeader bordersDoNotSurroundHeader1 = new BordersDoNotSurroundHeader();
                BordersDoNotSurroundFooter bordersDoNotSurroundFooter1 = new BordersDoNotSurroundFooter();
                ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
                DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 480 };
                DrawingGridHorizontalSpacing drawingGridHorizontalSpacing1 = new DrawingGridHorizontalSpacing() { Val = "120" };
                DisplayHorizontalDrawingGrid displayHorizontalDrawingGrid1 = new DisplayHorizontalDrawingGrid() { Val = 0 };
                DisplayVerticalDrawingGrid displayVerticalDrawingGrid1 = new DisplayVerticalDrawingGrid() { Val = 2 };
                CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.CompressPunctuation };

                HeaderShapeDefaults headerShapeDefaults1 = new HeaderShapeDefaults();
                Ovml.ShapeDefaults shapeDefaults1 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2050 };

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
                CompatibilitySetting compatibilitySetting1 = new CompatibilitySetting() { Name = CompatSettingNameValues.CompatibilityMode, Uri = "http://schemas.microsoft.com/office/word", Val = "15" };
                CompatibilitySetting compatibilitySetting2 = new CompatibilitySetting() { Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
                CompatibilitySetting compatibilitySetting3 = new CompatibilitySetting() { Name = CompatSettingNameValues.EnableOpenTypeFeatures, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
                CompatibilitySetting compatibilitySetting4 = new CompatibilitySetting() { Name = CompatSettingNameValues.DoNotFlipMirrorIndents, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
                CompatibilitySetting compatibilitySetting5 = new CompatibilitySetting() { Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };
                CompatibilitySetting compatibilitySetting6 = new CompatibilitySetting() { Name = new EnumValue<CompatSettingNameValues>() { InnerText = "useWord2013TrackBottomHyphenation" }, Uri = "http://schemas.microsoft.com/office/word", Val = "1" };

                compatibility1.Append(spaceForUnderline1);
                compatibility1.Append(balanceSingleByteDoubleByteWidth1);
                compatibility1.Append(doNotLeaveBackslashAlone1);
                compatibility1.Append(underlineTrailingSpaces1);
                compatibility1.Append(doNotExpandShiftReturn1);
                compatibility1.Append(adjustLineHeightInTable1);
                compatibility1.Append(useFarEastLayout1);
                compatibility1.Append(compatibilitySetting1);
                compatibility1.Append(compatibilitySetting2);
                compatibility1.Append(compatibilitySetting3);
                compatibility1.Append(compatibilitySetting4);
                compatibility1.Append(compatibilitySetting5);
                compatibility1.Append(compatibilitySetting6);

                Rsids rsids1 = new Rsids();
                RsidRoot rsidRoot1 = new RsidRoot() { Val = "00CA3D7A" };
                Rsid rsid1 = new Rsid() { Val = "000054B4" };
                Rsid rsid2 = new Rsid() { Val = "00065B4E" };
                Rsid rsid3 = new Rsid() { Val = "0007635F" };
                Rsid rsid4 = new Rsid() { Val = "00080C65" };
                Rsid rsid5 = new Rsid() { Val = "000941A0" };
                Rsid rsid6 = new Rsid() { Val = "000F6422" };
                Rsid rsid7 = new Rsid() { Val = "001012F0" };
                Rsid rsid8 = new Rsid() { Val = "00103EA3" };
                Rsid rsid9 = new Rsid() { Val = "00104B80" };
                Rsid rsid10 = new Rsid() { Val = "001346F0" };
                Rsid rsid11 = new Rsid() { Val = "0016775F" };
                Rsid rsid12 = new Rsid() { Val = "0018619C" };
                Rsid rsid13 = new Rsid() { Val = "00194B05" };
                Rsid rsid14 = new Rsid() { Val = "001A4243" };
                Rsid rsid15 = new Rsid() { Val = "001A4911" };
                Rsid rsid16 = new Rsid() { Val = "001C04A9" };
                Rsid rsid17 = new Rsid() { Val = "001D505C" };
                Rsid rsid18 = new Rsid() { Val = "00202379" };
                Rsid rsid19 = new Rsid() { Val = "002845DF" };
                Rsid rsid20 = new Rsid() { Val = "00292721" };
                Rsid rsid21 = new Rsid() { Val = "002B6D10" };
                Rsid rsid22 = new Rsid() { Val = "002C622B" };
                Rsid rsid23 = new Rsid() { Val = "003171DE" };
                Rsid rsid24 = new Rsid() { Val = "00355DCF" };
                Rsid rsid25 = new Rsid() { Val = "00365B1F" };
                Rsid rsid26 = new Rsid() { Val = "00397F67" };
                Rsid rsid27 = new Rsid() { Val = "003F6405" };
                Rsid rsid28 = new Rsid() { Val = "004207C4" };
                Rsid rsid29 = new Rsid() { Val = "004B57A5" };
                Rsid rsid30 = new Rsid() { Val = "004C13B3" };
                Rsid rsid31 = new Rsid() { Val = "004C7BBC" };
                Rsid rsid32 = new Rsid() { Val = "004D1D28" };
                Rsid rsid33 = new Rsid() { Val = "004E3417" };
                Rsid rsid34 = new Rsid() { Val = "0050595B" };
                Rsid rsid35 = new Rsid() { Val = "005523A1" };
                Rsid rsid36 = new Rsid() { Val = "00562A4C" };
                Rsid rsid37 = new Rsid() { Val = "00566146" };
                Rsid rsid38 = new Rsid() { Val = "00581081" };
                Rsid rsid39 = new Rsid() { Val = "005D21DA" };
                Rsid rsid40 = new Rsid() { Val = "005E3BCA" };
                Rsid rsid41 = new Rsid() { Val = "005F3C99" };
                Rsid rsid42 = new Rsid() { Val = "00606990" };
                Rsid rsid43 = new Rsid() { Val = "00610F0B" };
                Rsid rsid44 = new Rsid() { Val = "0062105A" };
                Rsid rsid45 = new Rsid() { Val = "006264D0" };
                Rsid rsid46 = new Rsid() { Val = "006723A1" };
                Rsid rsid47 = new Rsid() { Val = "00691BB8" };
                Rsid rsid48 = new Rsid() { Val = "006A2278" };
                Rsid rsid49 = new Rsid() { Val = "006B4156" };
                Rsid rsid50 = new Rsid() { Val = "007019E4" };
                Rsid rsid51 = new Rsid() { Val = "007607FB" };
                Rsid rsid52 = new Rsid() { Val = "00765DBF" };
                Rsid rsid53 = new Rsid() { Val = "007E1C13" };
                Rsid rsid54 = new Rsid() { Val = "00824D87" };
                Rsid rsid55 = new Rsid() { Val = "00855F24" };
                Rsid rsid56 = new Rsid() { Val = "00862A0C" };
                Rsid rsid57 = new Rsid() { Val = "0086389F" };
                Rsid rsid58 = new Rsid() { Val = "00882080" };
                Rsid rsid59 = new Rsid() { Val = "008C1D59" };
                Rsid rsid60 = new Rsid() { Val = "008C6DE4" };
                Rsid rsid61 = new Rsid() { Val = "00915C2D" };
                Rsid rsid62 = new Rsid() { Val = "00925EB7" };
                Rsid rsid63 = new Rsid() { Val = "009E02EC" };
                Rsid rsid64 = new Rsid() { Val = "009E4187" };
                Rsid rsid65 = new Rsid() { Val = "009E7336" };
                Rsid rsid66 = new Rsid() { Val = "009F6534" };
                Rsid rsid67 = new Rsid() { Val = "00A02E9E" };
                Rsid rsid68 = new Rsid() { Val = "00A578FC" };
                Rsid rsid69 = new Rsid() { Val = "00A6078C" };
                Rsid rsid70 = new Rsid() { Val = "00A6400D" };
                Rsid rsid71 = new Rsid() { Val = "00A66095" };
                Rsid rsid72 = new Rsid() { Val = "00A66443" };
                Rsid rsid73 = new Rsid() { Val = "00A713A7" };
                Rsid rsid74 = new Rsid() { Val = "00A75D4D" };
                Rsid rsid75 = new Rsid() { Val = "00A86FCD" };
                Rsid rsid76 = new Rsid() { Val = "00B26130" };
                Rsid rsid77 = new Rsid() { Val = "00B91D38" };
                Rsid rsid78 = new Rsid() { Val = "00B9778C" };
                Rsid rsid79 = new Rsid() { Val = "00BA4D5D" };
                Rsid rsid80 = new Rsid() { Val = "00BB5471" };
                Rsid rsid81 = new Rsid() { Val = "00C508FB" };
                Rsid rsid82 = new Rsid() { Val = "00C61C0F" };
                Rsid rsid83 = new Rsid() { Val = "00C747E9" };
                Rsid rsid84 = new Rsid() { Val = "00CA3D7A" };
                Rsid rsid85 = new Rsid() { Val = "00CB3FFE" };
                Rsid rsid86 = new Rsid() { Val = "00CC5221" };
                Rsid rsid87 = new Rsid() { Val = "00CF312B" };
                Rsid rsid88 = new Rsid() { Val = "00D43236" };
                Rsid rsid89 = new Rsid() { Val = "00D5039B" };
                Rsid rsid90 = new Rsid() { Val = "00D66A62" };
                Rsid rsid91 = new Rsid() { Val = "00E039AB" };
                Rsid rsid92 = new Rsid() { Val = "00E21A92" };
                Rsid rsid93 = new Rsid() { Val = "00E33911" };
                Rsid rsid94 = new Rsid() { Val = "00E34308" };
                Rsid rsid95 = new Rsid() { Val = "00E52DA5" };
                Rsid rsid96 = new Rsid() { Val = "00E76DB3" };
                Rsid rsid97 = new Rsid() { Val = "00EB24C7" };
                Rsid rsid98 = new Rsid() { Val = "00EB7726" };
                Rsid rsid99 = new Rsid() { Val = "00EC1462" };
                Rsid rsid100 = new Rsid() { Val = "00EC43E9" };
                Rsid rsid101 = new Rsid() { Val = "00ED2FDE" };
                Rsid rsid102 = new Rsid() { Val = "00EE6678" };
                Rsid rsid103 = new Rsid() { Val = "00F5167A" };
                Rsid rsid104 = new Rsid() { Val = "00F524DE" };
                Rsid rsid105 = new Rsid() { Val = "00F73903" };
                Rsid rsid106 = new Rsid() { Val = "00F84F36" };
                Rsid rsid107 = new Rsid() { Val = "00F90943" };
                Rsid rsid108 = new Rsid() { Val = "00FA6CF5" };
                Rsid rsid109 = new Rsid() { Val = "00FA79CC" };
                Rsid rsid110 = new Rsid() { Val = "00FC305E" };
                Rsid rsid111 = new Rsid() { Val = "1FB659BE" };
                Rsid rsid112 = new Rsid() { Val = "4CC70866" };
                Rsid rsid113 = new Rsid() { Val = "719D8CD9" };

                rsids1.Append(rsidRoot1);
                rsids1.Append(rsid1);
                rsids1.Append(rsid2);
                rsids1.Append(rsid3);
                rsids1.Append(rsid4);
                rsids1.Append(rsid5);
                rsids1.Append(rsid6);
                rsids1.Append(rsid7);
                rsids1.Append(rsid8);
                rsids1.Append(rsid9);
                rsids1.Append(rsid10);
                rsids1.Append(rsid11);
                rsids1.Append(rsid12);
                rsids1.Append(rsid13);
                rsids1.Append(rsid14);
                rsids1.Append(rsid15);
                rsids1.Append(rsid16);
                rsids1.Append(rsid17);
                rsids1.Append(rsid18);
                rsids1.Append(rsid19);
                rsids1.Append(rsid20);
                rsids1.Append(rsid21);
                rsids1.Append(rsid22);
                rsids1.Append(rsid23);
                rsids1.Append(rsid24);
                rsids1.Append(rsid25);
                rsids1.Append(rsid26);
                rsids1.Append(rsid27);
                rsids1.Append(rsid28);
                rsids1.Append(rsid29);
                rsids1.Append(rsid30);
                rsids1.Append(rsid31);
                rsids1.Append(rsid32);
                rsids1.Append(rsid33);
                rsids1.Append(rsid34);
                rsids1.Append(rsid35);
                rsids1.Append(rsid36);
                rsids1.Append(rsid37);
                rsids1.Append(rsid38);
                rsids1.Append(rsid39);
                rsids1.Append(rsid40);
                rsids1.Append(rsid41);
                rsids1.Append(rsid42);
                rsids1.Append(rsid43);
                rsids1.Append(rsid44);
                rsids1.Append(rsid45);
                rsids1.Append(rsid46);
                rsids1.Append(rsid47);
                rsids1.Append(rsid48);
                rsids1.Append(rsid49);
                rsids1.Append(rsid50);
                rsids1.Append(rsid51);
                rsids1.Append(rsid52);
                rsids1.Append(rsid53);
                rsids1.Append(rsid54);
                rsids1.Append(rsid55);
                rsids1.Append(rsid56);
                rsids1.Append(rsid57);
                rsids1.Append(rsid58);
                rsids1.Append(rsid59);
                rsids1.Append(rsid60);
                rsids1.Append(rsid61);
                rsids1.Append(rsid62);
                rsids1.Append(rsid63);
                rsids1.Append(rsid64);
                rsids1.Append(rsid65);
                rsids1.Append(rsid66);
                rsids1.Append(rsid67);
                rsids1.Append(rsid68);
                rsids1.Append(rsid69);
                rsids1.Append(rsid70);
                rsids1.Append(rsid71);
                rsids1.Append(rsid72);
                rsids1.Append(rsid73);
                rsids1.Append(rsid74);
                rsids1.Append(rsid75);
                rsids1.Append(rsid76);
                rsids1.Append(rsid77);
                rsids1.Append(rsid78);
                rsids1.Append(rsid79);
                rsids1.Append(rsid80);
                rsids1.Append(rsid81);
                rsids1.Append(rsid82);
                rsids1.Append(rsid83);
                rsids1.Append(rsid84);
                rsids1.Append(rsid85);
                rsids1.Append(rsid86);
                rsids1.Append(rsid87);
                rsids1.Append(rsid88);
                rsids1.Append(rsid89);
                rsids1.Append(rsid90);
                rsids1.Append(rsid91);
                rsids1.Append(rsid92);
                rsids1.Append(rsid93);
                rsids1.Append(rsid94);
                rsids1.Append(rsid95);
                rsids1.Append(rsid96);
                rsids1.Append(rsid97);
                rsids1.Append(rsid98);
                rsids1.Append(rsid99);
                rsids1.Append(rsid100);
                rsids1.Append(rsid101);
                rsids1.Append(rsid102);
                rsids1.Append(rsid103);
                rsids1.Append(rsid104);
                rsids1.Append(rsid105);
                rsids1.Append(rsid106);
                rsids1.Append(rsid107);
                rsids1.Append(rsid108);
                rsids1.Append(rsid109);
                rsids1.Append(rsid110);
                rsids1.Append(rsid111);
                rsids1.Append(rsid112);
                rsids1.Append(rsid113);

                M.MathProperties mathProperties1 = new M.MathProperties();
                M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
                M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
                M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
                M.SmallFraction smallFraction1 = new M.SmallFraction() { Val = M.BooleanValues.Zero };
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

                ShapeDefaults shapeDefaults2 = new ShapeDefaults();
                Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2050 };

                Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
                Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "2" };

                shapeLayout1.Append(shapeIdMap1);

                shapeDefaults2.Append(shapeDefaults3);
                shapeDefaults2.Append(shapeLayout1);
                DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
                ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
                W14.DocumentId documentId1 = new W14.DocumentId() { Val = "2C33AC77" };
                W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{AD2B8421-5027-4C76-BB91-45C383D0404A}" };

                settings1.Append(zoom1);
                settings1.Append(bordersDoNotSurroundHeader1);
                settings1.Append(bordersDoNotSurroundFooter1);
                settings1.Append(proofState1);
                settings1.Append(defaultTabStop1);
                settings1.Append(drawingGridHorizontalSpacing1);
                settings1.Append(displayHorizontalDrawingGrid1);
                settings1.Append(displayVerticalDrawingGrid1);
                settings1.Append(characterSpacingControl1);
                settings1.Append(headerShapeDefaults1);
                settings1.Append(footnoteDocumentWideProperties1);
                settings1.Append(endnoteDocumentWideProperties1);
                settings1.Append(compatibility1);
                settings1.Append(rsids1);
                settings1.Append(mathProperties1);
                settings1.Append(themeFontLanguages1);
                settings1.Append(colorSchemeMapping1);
                settings1.Append(shapeDefaults2);
                settings1.Append(decimalSymbol1);
                settings1.Append(listSeparator1);
                settings1.Append(documentId1);
                settings1.Append(persistentDocumentId1);

                documentSettingsPart1.Settings = settings1;
            }

            // Generates content of fontTablePart1.
            private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
            {
                Fonts fonts1 = new Fonts() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
                fonts1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                fonts1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                fonts1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                fonts1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                fonts1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                fonts1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                fonts1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                fonts1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

                Font font1 = new Font() { Name = "Wingdings" };
                Panose1Number panose1Number1 = new Panose1Number() { Val = "05000000000000000000" };
                FontCharSet fontCharSet1 = new FontCharSet() { Val = "02" };
                FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Auto };
                Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
                FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

                font1.Append(panose1Number1);
                font1.Append(fontCharSet1);
                font1.Append(fontFamily1);
                font1.Append(pitch1);
                font1.Append(fontSignature1);

                Font font2 = new Font() { Name = "Times New Roman" };
                Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
                FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
                FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
                Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
                FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

                font2.Append(panose1Number2);
                font2.Append(fontCharSet2);
                font2.Append(fontFamily2);
                font2.Append(pitch2);
                font2.Append(fontSignature2);

                Font font3 = new Font() { Name = "Calibri" };
                Panose1Number panose1Number3 = new Panose1Number() { Val = "020F0502020204030204" };
                FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
                FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
                Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
                FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C200247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

                font3.Append(panose1Number3);
                font3.Append(fontCharSet3);
                font3.Append(fontFamily3);
                font3.Append(pitch3);
                font3.Append(fontSignature3);

                Font font4 = new Font() { Name = "新細明體" };
                AltName altName1 = new AltName() { Val = "PMingLiU" };
                Panose1Number panose1Number4 = new Panose1Number() { Val = "02020500000000000000" };
                FontCharSet fontCharSet4 = new FontCharSet() { Val = "88" };
                FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
                Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
                FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "A00002FF", UnicodeSignature1 = "28CFFCFA", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00100001", CodePageSignature1 = "00000000" };

                font4.Append(altName1);
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
                FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E00006FF", UnicodeSignature1 = "420024FF", UnicodeSignature2 = "02000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

                font5.Append(panose1Number5);
                font5.Append(fontCharSet5);
                font5.Append(fontFamily5);
                font5.Append(pitch5);
                font5.Append(fontSignature5);

                Font font6 = new Font() { Name = "標楷體" };
                Panose1Number panose1Number6 = new Panose1Number() { Val = "03000509000000000000" };
                FontCharSet fontCharSet6 = new FontCharSet() { Val = "88" };
                FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Script };
                Pitch pitch6 = new Pitch() { Val = FontPitchValues.Fixed };
                FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "080E0000", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00100001", CodePageSignature1 = "00000000" };

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

            // Generates content of customXmlPart2.
            private void GenerateCustomXmlPart2Content(CustomXmlPart customXmlPart2)
            {
                System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart2.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
                writer.WriteRaw("<?mso-contentType?><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>DocumentLibraryForm</Display><Edit>DocumentLibraryForm</Edit><New>DocumentLibraryForm</New></FormTemplates>");
                writer.Flush();
                writer.Close();
            }

            // Generates content of customXmlPropertiesPart2.
            private void GenerateCustomXmlPropertiesPart2Content(CustomXmlPropertiesPart customXmlPropertiesPart2)
            {
                Ds.DataStoreItem dataStoreItem2 = new Ds.DataStoreItem() { ItemId = "{2DBB6213-F677-442A-8F39-A90EE0F1E055}" };
                dataStoreItem2.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

                Ds.SchemaReferences schemaReferences2 = new Ds.SchemaReferences();
                Ds.SchemaReference schemaReference3 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/sharepoint/v3/contenttype/forms" };

                schemaReferences2.Append(schemaReference3);

                dataStoreItem2.Append(schemaReferences2);

                customXmlPropertiesPart2.DataStoreItem = dataStoreItem2;
            }

            // Generates content of customXmlPart3.
            private void GenerateCustomXmlPart3Content(CustomXmlPart customXmlPart3)
            {
                System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart3.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><ct:contentTypeSchema ct:_=\"\" ma:_=\"\" ma:contentTypeName=\"文件\" ma:contentTypeID=\"0x010100D465D9393F8B7B429C8C9CE9DCEB2AA6\" ma:contentTypeVersion=\"3\" ma:contentTypeDescription=\"建立新的文件。\" ma:contentTypeScope=\"\" ma:versionID=\"0378f4f1ff7daa522c0e5e98de0b0a1f\" xmlns:ct=\"http://schemas.microsoft.com/office/2006/metadata/contentType\" xmlns:ma=\"http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes\">\r\n<xsd:schema targetNamespace=\"http://schemas.microsoft.com/office/2006/metadata/properties\" ma:root=\"true\" ma:fieldsID=\"378dabfd38230a78222a5f01fcd449b0\" ns2:_=\"\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" xmlns:ns2=\"020e0566-15c6-45f6-927c-a14f1387091e\">\r\n<xsd:import namespace=\"020e0566-15c6-45f6-927c-a14f1387091e\"/>\r\n<xsd:element name=\"properties\">\r\n<xsd:complexType>\r\n<xsd:sequence>\r\n<xsd:element name=\"documentManagement\">\r\n<xsd:complexType>\r\n<xsd:all>\r\n<xsd:element ref=\"ns2:MediaServiceMetadata\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceFastMetadata\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceObjectDetectorVersions\" minOccurs=\"0\"/>\r\n</xsd:all>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:sequence>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"020e0566-15c6-45f6-927c-a14f1387091e\" elementFormDefault=\"qualified\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:dms=\"http://schemas.microsoft.com/office/2006/documentManagement/types\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\">\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/2006/documentManagement/types\"/>\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"/>\r\n<xsd:element name=\"MediaServiceMetadata\" ma:index=\"8\" nillable=\"true\" ma:displayName=\"MediaServiceMetadata\" ma:hidden=\"true\" ma:internalName=\"MediaServiceMetadata\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceFastMetadata\" ma:index=\"9\" nillable=\"true\" ma:displayName=\"MediaServiceFastMetadata\" ma:hidden=\"true\" ma:internalName=\"MediaServiceFastMetadata\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceObjectDetectorVersions\" ma:index=\"10\" nillable=\"true\" ma:displayName=\"MediaServiceObjectDetectorVersions\" ma:hidden=\"true\" ma:indexed=\"true\" ma:internalName=\"MediaServiceObjectDetectorVersions\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" elementFormDefault=\"qualified\" attributeFormDefault=\"unqualified\" blockDefault=\"#all\" xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:odoc=\"http://schemas.microsoft.com/internal/obd\">\r\n<xsd:import namespace=\"http://purl.org/dc/elements/1.1/\" schemaLocation=\"http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd\"/>\r\n<xsd:import namespace=\"http://purl.org/dc/terms/\" schemaLocation=\"http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcterms.xsd\"/>\r\n<xsd:element name=\"coreProperties\" type=\"CT_coreProperties\"/>\r\n<xsd:complexType name=\"CT_coreProperties\">\r\n<xsd:all>\r\n<xsd:element ref=\"dc:creator\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dcterms:created\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dc:identifier\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"contentType\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\" ma:index=\"0\" ma:displayName=\"內容類型\"/>\r\n<xsd:element ref=\"dc:title\" minOccurs=\"0\" maxOccurs=\"1\" ma:index=\"4\" ma:displayName=\"標題\"/>\r\n<xsd:element ref=\"dc:subject\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dc:description\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"keywords\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element ref=\"dc:language\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"category\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element name=\"version\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element name=\"revision\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\">\r\n<xsd:annotation>\r\n<xsd:documentation>\r\n                        This value indicates the number of saves or revisions. The application is responsible for updating this value after each revision.\r\n                    </xsd:documentation>\r\n</xsd:annotation>\r\n</xsd:element>\r\n<xsd:element name=\"lastModifiedBy\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element ref=\"dcterms:modified\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"contentStatus\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n</xsd:all>\r\n</xsd:complexType>\r\n</xsd:schema>\r\n<xs:schema targetNamespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\" elementFormDefault=\"qualified\" attributeFormDefault=\"unqualified\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">\r\n<xs:element name=\"Person\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:DisplayName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:AccountId\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:AccountType\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"DisplayName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"AccountId\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"AccountType\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"BDCAssociatedEntity\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:BDCEntity\" minOccurs=\"0\" maxOccurs=\"unbounded\"></xs:element>\r\n</xs:sequence>\r\n<xs:attribute ref=\"pc:EntityNamespace\"></xs:attribute>\r\n<xs:attribute ref=\"pc:EntityName\"></xs:attribute>\r\n<xs:attribute ref=\"pc:SystemInstanceName\"></xs:attribute>\r\n<xs:attribute ref=\"pc:AssociationName\"></xs:attribute>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:attribute name=\"EntityNamespace\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"EntityName\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"SystemInstanceName\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"AssociationName\" type=\"xs:string\"></xs:attribute>\r\n<xs:element name=\"BDCEntity\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:EntityDisplayName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityInstanceReference\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId1\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId2\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId3\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId4\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId5\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"EntityDisplayName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityInstanceReference\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId1\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId2\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId3\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId4\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId5\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"Terms\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:TermInfo\" minOccurs=\"0\" maxOccurs=\"unbounded\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"TermInfo\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:TermName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:TermId\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"TermName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"TermId\" type=\"xs:string\"></xs:element>\r\n</xs:schema>\r\n</ct:contentTypeSchema>");
                writer.Flush();
                writer.Close();
            }

            // Generates content of customXmlPropertiesPart3.
            private void GenerateCustomXmlPropertiesPart3Content(CustomXmlPropertiesPart customXmlPropertiesPart3)
            {
                Ds.DataStoreItem dataStoreItem3 = new Ds.DataStoreItem() { ItemId = "{135486CB-4DBB-49C7-85AD-66277254F5C7}" };
                dataStoreItem3.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

                Ds.SchemaReferences schemaReferences3 = new Ds.SchemaReferences();
                Ds.SchemaReference schemaReference4 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/contentType" };
                Ds.SchemaReference schemaReference5 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes" };
                Ds.SchemaReference schemaReference6 = new Ds.SchemaReference() { Uri = "http://www.w3.org/2001/XMLSchema" };
                Ds.SchemaReference schemaReference7 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties" };
                Ds.SchemaReference schemaReference8 = new Ds.SchemaReference() { Uri = "020e0566-15c6-45f6-927c-a14f1387091e" };
                Ds.SchemaReference schemaReference9 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/documentManagement/types" };
                Ds.SchemaReference schemaReference10 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls" };
                Ds.SchemaReference schemaReference11 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" };
                Ds.SchemaReference schemaReference12 = new Ds.SchemaReference() { Uri = "http://purl.org/dc/elements/1.1/" };
                Ds.SchemaReference schemaReference13 = new Ds.SchemaReference() { Uri = "http://purl.org/dc/terms/" };
                Ds.SchemaReference schemaReference14 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/internal/obd" };

                schemaReferences3.Append(schemaReference4);
                schemaReferences3.Append(schemaReference5);
                schemaReferences3.Append(schemaReference6);
                schemaReferences3.Append(schemaReference7);
                schemaReferences3.Append(schemaReference8);
                schemaReferences3.Append(schemaReference9);
                schemaReferences3.Append(schemaReference10);
                schemaReferences3.Append(schemaReference11);
                schemaReferences3.Append(schemaReference12);
                schemaReferences3.Append(schemaReference13);
                schemaReferences3.Append(schemaReference14);

                dataStoreItem3.Append(schemaReferences3);

                customXmlPropertiesPart3.DataStoreItem = dataStoreItem3;
            }

            // Generates content of styleDefinitionsPart1.
            private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
            {
                Styles styles1 = new Styles() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh" } };
                styles1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                styles1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                styles1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                styles1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                styles1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                styles1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                styles1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                styles1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");

                DocDefaults docDefaults1 = new DocDefaults();

                RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

                RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
                RunFonts runFonts71 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi };
                Kern kern30 = new Kern() { Val = (UInt32Value)2U };
                FontSize fontSize63 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "22" };
                Languages languages27 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

                runPropertiesBaseStyle1.Append(runFonts71);
                runPropertiesBaseStyle1.Append(kern30);
                runPropertiesBaseStyle1.Append(fontSize63);
                runPropertiesBaseStyle1.Append(fontSizeComplexScript64);
                runPropertiesBaseStyle1.Append(languages27);

                runPropertiesDefault1.Append(runPropertiesBaseStyle1);
                ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

                docDefaults1.Append(runPropertiesDefault1);
                docDefaults1.Append(paragraphPropertiesDefault1);

                LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 376 };
                LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "index 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "index 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "index 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "index 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "index 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "index 6", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "index 7", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "index 8", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "index 9", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Normal Indent", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "footnote text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "annotation text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "header", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "footer", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "index heading", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "table of figures", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "envelope address", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "envelope return", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "footnote reference", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "annotation reference", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "line number", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "page number", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "endnote reference", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "endnote text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "table of authorities", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "macro", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "toa heading", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "List", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "List Bullet", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Number", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "List 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "List 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "List 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "List 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "List Bullet 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "List Bullet 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "List Bullet 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "List Bullet 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "List Number 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "List Number 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "List Number 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "List Number 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Closing", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Signature", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Body Text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Body Text Indent", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "List Continue", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "List Continue 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "List Continue 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "List Continue 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "List Continue 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Message Header", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Salutation", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Date", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Body Text First Indent 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Note Heading", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Body Text 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Body Text 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Body Text Indent 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Block Text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "FollowedHyperlink", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Document Map", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Plain Text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "E-mail Signature", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "HTML Top of Form", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "HTML Bottom of Form", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Normal (Web)", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "HTML Acronym", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "HTML Address", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "HTML Cite", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "HTML Code", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "HTML Definition", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "HTML Keyboard", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "HTML Preformatted", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "HTML Sample", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "HTML Typewriter", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "HTML Variable", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "annotation subject", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "No List", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Outline List 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Outline List 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Outline List 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Balloon Text", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 59 };
                LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", SemiHidden = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Revision", SemiHidden = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
                LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
                LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
                LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
                LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
                LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
                LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
                LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
                LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
                LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
                LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
                LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
                LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
                LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
                LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
                LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
                LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
                LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
                LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
                LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
                LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo282 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo283 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo284 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo285 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo286 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo287 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo288 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo289 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo290 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo291 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo292 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo293 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo294 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo295 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo296 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo297 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo298 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo299 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo300 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo301 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo302 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo303 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo304 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo305 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo306 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo307 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo308 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo309 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo310 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo311 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo312 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo313 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo314 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo315 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo316 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo317 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo318 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo319 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo320 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo321 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo322 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo323 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo324 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo325 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo326 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo327 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo328 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo329 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo330 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo331 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo332 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo333 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo334 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo335 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo336 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo337 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo338 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo339 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo340 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo341 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo342 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo343 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo344 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo345 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo346 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo347 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo348 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo349 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo350 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo351 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo352 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo353 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo354 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo355 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo356 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo357 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo358 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo359 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo360 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo361 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo362 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo363 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo364 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo365 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
                LatentStyleExceptionInfo latentStyleExceptionInfo366 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
                LatentStyleExceptionInfo latentStyleExceptionInfo367 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
                LatentStyleExceptionInfo latentStyleExceptionInfo368 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
                LatentStyleExceptionInfo latentStyleExceptionInfo369 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
                LatentStyleExceptionInfo latentStyleExceptionInfo370 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
                LatentStyleExceptionInfo latentStyleExceptionInfo371 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };
                LatentStyleExceptionInfo latentStyleExceptionInfo372 = new LatentStyleExceptionInfo() { Name = "Mention", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo373 = new LatentStyleExceptionInfo() { Name = "Smart Hyperlink", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo374 = new LatentStyleExceptionInfo() { Name = "Hashtag", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo375 = new LatentStyleExceptionInfo() { Name = "Unresolved Mention", SemiHidden = true, UnhideWhenUsed = true };
                LatentStyleExceptionInfo latentStyleExceptionInfo376 = new LatentStyleExceptionInfo() { Name = "Smart Link", SemiHidden = true, UnhideWhenUsed = true };

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
                latentStyles1.Append(latentStyleExceptionInfo138);
                latentStyles1.Append(latentStyleExceptionInfo139);
                latentStyles1.Append(latentStyleExceptionInfo140);
                latentStyles1.Append(latentStyleExceptionInfo141);
                latentStyles1.Append(latentStyleExceptionInfo142);
                latentStyles1.Append(latentStyleExceptionInfo143);
                latentStyles1.Append(latentStyleExceptionInfo144);
                latentStyles1.Append(latentStyleExceptionInfo145);
                latentStyles1.Append(latentStyleExceptionInfo146);
                latentStyles1.Append(latentStyleExceptionInfo147);
                latentStyles1.Append(latentStyleExceptionInfo148);
                latentStyles1.Append(latentStyleExceptionInfo149);
                latentStyles1.Append(latentStyleExceptionInfo150);
                latentStyles1.Append(latentStyleExceptionInfo151);
                latentStyles1.Append(latentStyleExceptionInfo152);
                latentStyles1.Append(latentStyleExceptionInfo153);
                latentStyles1.Append(latentStyleExceptionInfo154);
                latentStyles1.Append(latentStyleExceptionInfo155);
                latentStyles1.Append(latentStyleExceptionInfo156);
                latentStyles1.Append(latentStyleExceptionInfo157);
                latentStyles1.Append(latentStyleExceptionInfo158);
                latentStyles1.Append(latentStyleExceptionInfo159);
                latentStyles1.Append(latentStyleExceptionInfo160);
                latentStyles1.Append(latentStyleExceptionInfo161);
                latentStyles1.Append(latentStyleExceptionInfo162);
                latentStyles1.Append(latentStyleExceptionInfo163);
                latentStyles1.Append(latentStyleExceptionInfo164);
                latentStyles1.Append(latentStyleExceptionInfo165);
                latentStyles1.Append(latentStyleExceptionInfo166);
                latentStyles1.Append(latentStyleExceptionInfo167);
                latentStyles1.Append(latentStyleExceptionInfo168);
                latentStyles1.Append(latentStyleExceptionInfo169);
                latentStyles1.Append(latentStyleExceptionInfo170);
                latentStyles1.Append(latentStyleExceptionInfo171);
                latentStyles1.Append(latentStyleExceptionInfo172);
                latentStyles1.Append(latentStyleExceptionInfo173);
                latentStyles1.Append(latentStyleExceptionInfo174);
                latentStyles1.Append(latentStyleExceptionInfo175);
                latentStyles1.Append(latentStyleExceptionInfo176);
                latentStyles1.Append(latentStyleExceptionInfo177);
                latentStyles1.Append(latentStyleExceptionInfo178);
                latentStyles1.Append(latentStyleExceptionInfo179);
                latentStyles1.Append(latentStyleExceptionInfo180);
                latentStyles1.Append(latentStyleExceptionInfo181);
                latentStyles1.Append(latentStyleExceptionInfo182);
                latentStyles1.Append(latentStyleExceptionInfo183);
                latentStyles1.Append(latentStyleExceptionInfo184);
                latentStyles1.Append(latentStyleExceptionInfo185);
                latentStyles1.Append(latentStyleExceptionInfo186);
                latentStyles1.Append(latentStyleExceptionInfo187);
                latentStyles1.Append(latentStyleExceptionInfo188);
                latentStyles1.Append(latentStyleExceptionInfo189);
                latentStyles1.Append(latentStyleExceptionInfo190);
                latentStyles1.Append(latentStyleExceptionInfo191);
                latentStyles1.Append(latentStyleExceptionInfo192);
                latentStyles1.Append(latentStyleExceptionInfo193);
                latentStyles1.Append(latentStyleExceptionInfo194);
                latentStyles1.Append(latentStyleExceptionInfo195);
                latentStyles1.Append(latentStyleExceptionInfo196);
                latentStyles1.Append(latentStyleExceptionInfo197);
                latentStyles1.Append(latentStyleExceptionInfo198);
                latentStyles1.Append(latentStyleExceptionInfo199);
                latentStyles1.Append(latentStyleExceptionInfo200);
                latentStyles1.Append(latentStyleExceptionInfo201);
                latentStyles1.Append(latentStyleExceptionInfo202);
                latentStyles1.Append(latentStyleExceptionInfo203);
                latentStyles1.Append(latentStyleExceptionInfo204);
                latentStyles1.Append(latentStyleExceptionInfo205);
                latentStyles1.Append(latentStyleExceptionInfo206);
                latentStyles1.Append(latentStyleExceptionInfo207);
                latentStyles1.Append(latentStyleExceptionInfo208);
                latentStyles1.Append(latentStyleExceptionInfo209);
                latentStyles1.Append(latentStyleExceptionInfo210);
                latentStyles1.Append(latentStyleExceptionInfo211);
                latentStyles1.Append(latentStyleExceptionInfo212);
                latentStyles1.Append(latentStyleExceptionInfo213);
                latentStyles1.Append(latentStyleExceptionInfo214);
                latentStyles1.Append(latentStyleExceptionInfo215);
                latentStyles1.Append(latentStyleExceptionInfo216);
                latentStyles1.Append(latentStyleExceptionInfo217);
                latentStyles1.Append(latentStyleExceptionInfo218);
                latentStyles1.Append(latentStyleExceptionInfo219);
                latentStyles1.Append(latentStyleExceptionInfo220);
                latentStyles1.Append(latentStyleExceptionInfo221);
                latentStyles1.Append(latentStyleExceptionInfo222);
                latentStyles1.Append(latentStyleExceptionInfo223);
                latentStyles1.Append(latentStyleExceptionInfo224);
                latentStyles1.Append(latentStyleExceptionInfo225);
                latentStyles1.Append(latentStyleExceptionInfo226);
                latentStyles1.Append(latentStyleExceptionInfo227);
                latentStyles1.Append(latentStyleExceptionInfo228);
                latentStyles1.Append(latentStyleExceptionInfo229);
                latentStyles1.Append(latentStyleExceptionInfo230);
                latentStyles1.Append(latentStyleExceptionInfo231);
                latentStyles1.Append(latentStyleExceptionInfo232);
                latentStyles1.Append(latentStyleExceptionInfo233);
                latentStyles1.Append(latentStyleExceptionInfo234);
                latentStyles1.Append(latentStyleExceptionInfo235);
                latentStyles1.Append(latentStyleExceptionInfo236);
                latentStyles1.Append(latentStyleExceptionInfo237);
                latentStyles1.Append(latentStyleExceptionInfo238);
                latentStyles1.Append(latentStyleExceptionInfo239);
                latentStyles1.Append(latentStyleExceptionInfo240);
                latentStyles1.Append(latentStyleExceptionInfo241);
                latentStyles1.Append(latentStyleExceptionInfo242);
                latentStyles1.Append(latentStyleExceptionInfo243);
                latentStyles1.Append(latentStyleExceptionInfo244);
                latentStyles1.Append(latentStyleExceptionInfo245);
                latentStyles1.Append(latentStyleExceptionInfo246);
                latentStyles1.Append(latentStyleExceptionInfo247);
                latentStyles1.Append(latentStyleExceptionInfo248);
                latentStyles1.Append(latentStyleExceptionInfo249);
                latentStyles1.Append(latentStyleExceptionInfo250);
                latentStyles1.Append(latentStyleExceptionInfo251);
                latentStyles1.Append(latentStyleExceptionInfo252);
                latentStyles1.Append(latentStyleExceptionInfo253);
                latentStyles1.Append(latentStyleExceptionInfo254);
                latentStyles1.Append(latentStyleExceptionInfo255);
                latentStyles1.Append(latentStyleExceptionInfo256);
                latentStyles1.Append(latentStyleExceptionInfo257);
                latentStyles1.Append(latentStyleExceptionInfo258);
                latentStyles1.Append(latentStyleExceptionInfo259);
                latentStyles1.Append(latentStyleExceptionInfo260);
                latentStyles1.Append(latentStyleExceptionInfo261);
                latentStyles1.Append(latentStyleExceptionInfo262);
                latentStyles1.Append(latentStyleExceptionInfo263);
                latentStyles1.Append(latentStyleExceptionInfo264);
                latentStyles1.Append(latentStyleExceptionInfo265);
                latentStyles1.Append(latentStyleExceptionInfo266);
                latentStyles1.Append(latentStyleExceptionInfo267);
                latentStyles1.Append(latentStyleExceptionInfo268);
                latentStyles1.Append(latentStyleExceptionInfo269);
                latentStyles1.Append(latentStyleExceptionInfo270);
                latentStyles1.Append(latentStyleExceptionInfo271);
                latentStyles1.Append(latentStyleExceptionInfo272);
                latentStyles1.Append(latentStyleExceptionInfo273);
                latentStyles1.Append(latentStyleExceptionInfo274);
                latentStyles1.Append(latentStyleExceptionInfo275);
                latentStyles1.Append(latentStyleExceptionInfo276);
                latentStyles1.Append(latentStyleExceptionInfo277);
                latentStyles1.Append(latentStyleExceptionInfo278);
                latentStyles1.Append(latentStyleExceptionInfo279);
                latentStyles1.Append(latentStyleExceptionInfo280);
                latentStyles1.Append(latentStyleExceptionInfo281);
                latentStyles1.Append(latentStyleExceptionInfo282);
                latentStyles1.Append(latentStyleExceptionInfo283);
                latentStyles1.Append(latentStyleExceptionInfo284);
                latentStyles1.Append(latentStyleExceptionInfo285);
                latentStyles1.Append(latentStyleExceptionInfo286);
                latentStyles1.Append(latentStyleExceptionInfo287);
                latentStyles1.Append(latentStyleExceptionInfo288);
                latentStyles1.Append(latentStyleExceptionInfo289);
                latentStyles1.Append(latentStyleExceptionInfo290);
                latentStyles1.Append(latentStyleExceptionInfo291);
                latentStyles1.Append(latentStyleExceptionInfo292);
                latentStyles1.Append(latentStyleExceptionInfo293);
                latentStyles1.Append(latentStyleExceptionInfo294);
                latentStyles1.Append(latentStyleExceptionInfo295);
                latentStyles1.Append(latentStyleExceptionInfo296);
                latentStyles1.Append(latentStyleExceptionInfo297);
                latentStyles1.Append(latentStyleExceptionInfo298);
                latentStyles1.Append(latentStyleExceptionInfo299);
                latentStyles1.Append(latentStyleExceptionInfo300);
                latentStyles1.Append(latentStyleExceptionInfo301);
                latentStyles1.Append(latentStyleExceptionInfo302);
                latentStyles1.Append(latentStyleExceptionInfo303);
                latentStyles1.Append(latentStyleExceptionInfo304);
                latentStyles1.Append(latentStyleExceptionInfo305);
                latentStyles1.Append(latentStyleExceptionInfo306);
                latentStyles1.Append(latentStyleExceptionInfo307);
                latentStyles1.Append(latentStyleExceptionInfo308);
                latentStyles1.Append(latentStyleExceptionInfo309);
                latentStyles1.Append(latentStyleExceptionInfo310);
                latentStyles1.Append(latentStyleExceptionInfo311);
                latentStyles1.Append(latentStyleExceptionInfo312);
                latentStyles1.Append(latentStyleExceptionInfo313);
                latentStyles1.Append(latentStyleExceptionInfo314);
                latentStyles1.Append(latentStyleExceptionInfo315);
                latentStyles1.Append(latentStyleExceptionInfo316);
                latentStyles1.Append(latentStyleExceptionInfo317);
                latentStyles1.Append(latentStyleExceptionInfo318);
                latentStyles1.Append(latentStyleExceptionInfo319);
                latentStyles1.Append(latentStyleExceptionInfo320);
                latentStyles1.Append(latentStyleExceptionInfo321);
                latentStyles1.Append(latentStyleExceptionInfo322);
                latentStyles1.Append(latentStyleExceptionInfo323);
                latentStyles1.Append(latentStyleExceptionInfo324);
                latentStyles1.Append(latentStyleExceptionInfo325);
                latentStyles1.Append(latentStyleExceptionInfo326);
                latentStyles1.Append(latentStyleExceptionInfo327);
                latentStyles1.Append(latentStyleExceptionInfo328);
                latentStyles1.Append(latentStyleExceptionInfo329);
                latentStyles1.Append(latentStyleExceptionInfo330);
                latentStyles1.Append(latentStyleExceptionInfo331);
                latentStyles1.Append(latentStyleExceptionInfo332);
                latentStyles1.Append(latentStyleExceptionInfo333);
                latentStyles1.Append(latentStyleExceptionInfo334);
                latentStyles1.Append(latentStyleExceptionInfo335);
                latentStyles1.Append(latentStyleExceptionInfo336);
                latentStyles1.Append(latentStyleExceptionInfo337);
                latentStyles1.Append(latentStyleExceptionInfo338);
                latentStyles1.Append(latentStyleExceptionInfo339);
                latentStyles1.Append(latentStyleExceptionInfo340);
                latentStyles1.Append(latentStyleExceptionInfo341);
                latentStyles1.Append(latentStyleExceptionInfo342);
                latentStyles1.Append(latentStyleExceptionInfo343);
                latentStyles1.Append(latentStyleExceptionInfo344);
                latentStyles1.Append(latentStyleExceptionInfo345);
                latentStyles1.Append(latentStyleExceptionInfo346);
                latentStyles1.Append(latentStyleExceptionInfo347);
                latentStyles1.Append(latentStyleExceptionInfo348);
                latentStyles1.Append(latentStyleExceptionInfo349);
                latentStyles1.Append(latentStyleExceptionInfo350);
                latentStyles1.Append(latentStyleExceptionInfo351);
                latentStyles1.Append(latentStyleExceptionInfo352);
                latentStyles1.Append(latentStyleExceptionInfo353);
                latentStyles1.Append(latentStyleExceptionInfo354);
                latentStyles1.Append(latentStyleExceptionInfo355);
                latentStyles1.Append(latentStyleExceptionInfo356);
                latentStyles1.Append(latentStyleExceptionInfo357);
                latentStyles1.Append(latentStyleExceptionInfo358);
                latentStyles1.Append(latentStyleExceptionInfo359);
                latentStyles1.Append(latentStyleExceptionInfo360);
                latentStyles1.Append(latentStyleExceptionInfo361);
                latentStyles1.Append(latentStyleExceptionInfo362);
                latentStyles1.Append(latentStyleExceptionInfo363);
                latentStyles1.Append(latentStyleExceptionInfo364);
                latentStyles1.Append(latentStyleExceptionInfo365);
                latentStyles1.Append(latentStyleExceptionInfo366);
                latentStyles1.Append(latentStyleExceptionInfo367);
                latentStyles1.Append(latentStyleExceptionInfo368);
                latentStyles1.Append(latentStyleExceptionInfo369);
                latentStyles1.Append(latentStyleExceptionInfo370);
                latentStyles1.Append(latentStyleExceptionInfo371);
                latentStyles1.Append(latentStyleExceptionInfo372);
                latentStyles1.Append(latentStyleExceptionInfo373);
                latentStyles1.Append(latentStyleExceptionInfo374);
                latentStyles1.Append(latentStyleExceptionInfo375);
                latentStyles1.Append(latentStyleExceptionInfo376);

                Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
                StyleName styleName1 = new StyleName() { Val = "Normal" };
                PrimaryStyle primaryStyle1 = new PrimaryStyle();

                StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
                WidowControl widowControl1 = new WidowControl() { Val = false };

                styleParagraphProperties1.Append(widowControl1);

                style1.Append(styleName1);
                style1.Append(primaryStyle1);
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
                StyleName styleName5 = new StyleName() { Val = "Date" };
                BasedOn basedOn1 = new BasedOn() { Val = "a" };
                NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "a" };
                LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "a4" };
                UIPriority uIPriority4 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden4 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
                Rsid rsid114 = new Rsid() { Val = "00CA3D7A" };

                StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
                Justification justification6 = new Justification() { Val = JustificationValues.Right };

                styleParagraphProperties2.Append(justification6);

                style5.Append(styleName5);
                style5.Append(basedOn1);
                style5.Append(nextParagraphStyle1);
                style5.Append(linkedStyle1);
                style5.Append(uIPriority4);
                style5.Append(semiHidden4);
                style5.Append(unhideWhenUsed4);
                style5.Append(rsid114);
                style5.Append(styleParagraphProperties2);

                Style style6 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
                StyleName styleName6 = new StyleName() { Val = "日期 字元" };
                BasedOn basedOn2 = new BasedOn() { Val = "a0" };
                LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a3" };
                UIPriority uIPriority5 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden5 = new SemiHidden();
                Rsid rsid115 = new Rsid() { Val = "00CA3D7A" };

                style6.Append(styleName6);
                style6.Append(basedOn2);
                style6.Append(linkedStyle2);
                style6.Append(uIPriority5);
                style6.Append(semiHidden5);
                style6.Append(rsid115);

                Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5" };
                StyleName styleName7 = new StyleName() { Val = "header" };
                BasedOn basedOn3 = new BasedOn() { Val = "a" };
                LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "a6" };
                UIPriority uIPriority6 = new UIPriority() { Val = 99 };
                UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
                Rsid rsid116 = new Rsid() { Val = "00A66095" };

                StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

                Tabs tabs1 = new Tabs();
                TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
                TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

                tabs1.Append(tabStop1);
                tabs1.Append(tabStop2);
                SnapToGrid snapToGrid7 = new SnapToGrid() { Val = false };

                styleParagraphProperties3.Append(tabs1);
                styleParagraphProperties3.Append(snapToGrid7);

                StyleRunProperties styleRunProperties1 = new StyleRunProperties();
                FontSize fontSize64 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties1.Append(fontSize64);
                styleRunProperties1.Append(fontSizeComplexScript65);

                style7.Append(styleName7);
                style7.Append(basedOn3);
                style7.Append(linkedStyle3);
                style7.Append(uIPriority6);
                style7.Append(unhideWhenUsed5);
                style7.Append(rsid116);
                style7.Append(styleParagraphProperties3);
                style7.Append(styleRunProperties1);

                Style style8 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
                StyleName styleName8 = new StyleName() { Val = "頁首 字元" };
                BasedOn basedOn4 = new BasedOn() { Val = "a0" };
                LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a5" };
                UIPriority uIPriority7 = new UIPriority() { Val = 99 };
                Rsid rsid117 = new Rsid() { Val = "00A66095" };

                StyleRunProperties styleRunProperties2 = new StyleRunProperties();
                FontSize fontSize65 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties2.Append(fontSize65);
                styleRunProperties2.Append(fontSizeComplexScript66);

                style8.Append(styleName8);
                style8.Append(basedOn4);
                style8.Append(linkedStyle4);
                style8.Append(uIPriority7);
                style8.Append(rsid117);
                style8.Append(styleRunProperties2);

                Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "a7" };
                StyleName styleName9 = new StyleName() { Val = "footer" };
                BasedOn basedOn5 = new BasedOn() { Val = "a" };
                LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "a8" };
                UIPriority uIPriority8 = new UIPriority() { Val = 99 };
                UnhideWhenUsed unhideWhenUsed6 = new UnhideWhenUsed();
                Rsid rsid118 = new Rsid() { Val = "00A66095" };

                StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();

                Tabs tabs2 = new Tabs();
                TabStop tabStop3 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
                TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

                tabs2.Append(tabStop3);
                tabs2.Append(tabStop4);
                SnapToGrid snapToGrid8 = new SnapToGrid() { Val = false };

                styleParagraphProperties4.Append(tabs2);
                styleParagraphProperties4.Append(snapToGrid8);

                StyleRunProperties styleRunProperties3 = new StyleRunProperties();
                FontSize fontSize66 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties3.Append(fontSize66);
                styleRunProperties3.Append(fontSizeComplexScript67);

                style9.Append(styleName9);
                style9.Append(basedOn5);
                style9.Append(linkedStyle5);
                style9.Append(uIPriority8);
                style9.Append(unhideWhenUsed6);
                style9.Append(rsid118);
                style9.Append(styleParagraphProperties4);
                style9.Append(styleRunProperties3);

                Style style10 = new Style() { Type = StyleValues.Character, StyleId = "a8", CustomStyle = true };
                StyleName styleName10 = new StyleName() { Val = "頁尾 字元" };
                BasedOn basedOn6 = new BasedOn() { Val = "a0" };
                LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "a7" };
                UIPriority uIPriority9 = new UIPriority() { Val = 99 };
                Rsid rsid119 = new Rsid() { Val = "00A66095" };

                StyleRunProperties styleRunProperties4 = new StyleRunProperties();
                FontSize fontSize67 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties4.Append(fontSize67);
                styleRunProperties4.Append(fontSizeComplexScript68);

                style10.Append(styleName10);
                style10.Append(basedOn6);
                style10.Append(linkedStyle6);
                style10.Append(uIPriority9);
                style10.Append(rsid119);
                style10.Append(styleRunProperties4);

                Style style11 = new Style() { Type = StyleValues.Paragraph, StyleId = "a9" };
                StyleName styleName11 = new StyleName() { Val = "Balloon Text" };
                BasedOn basedOn7 = new BasedOn() { Val = "a" };
                LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "aa" };
                UIPriority uIPriority10 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden6 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed7 = new UnhideWhenUsed();
                Rsid rsid120 = new Rsid() { Val = "00765DBF" };

                StyleRunProperties styleRunProperties5 = new StyleRunProperties();
                RunFonts runFonts72 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                FontSize fontSize68 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "18" };

                styleRunProperties5.Append(runFonts72);
                styleRunProperties5.Append(fontSize68);
                styleRunProperties5.Append(fontSizeComplexScript69);

                style11.Append(styleName11);
                style11.Append(basedOn7);
                style11.Append(linkedStyle7);
                style11.Append(uIPriority10);
                style11.Append(semiHidden6);
                style11.Append(unhideWhenUsed7);
                style11.Append(rsid120);
                style11.Append(styleRunProperties5);

                Style style12 = new Style() { Type = StyleValues.Character, StyleId = "aa", CustomStyle = true };
                StyleName styleName12 = new StyleName() { Val = "註解方塊文字 字元" };
                BasedOn basedOn8 = new BasedOn() { Val = "a0" };
                LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "a9" };
                UIPriority uIPriority11 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden7 = new SemiHidden();
                Rsid rsid121 = new Rsid() { Val = "00765DBF" };

                StyleRunProperties styleRunProperties6 = new StyleRunProperties();
                RunFonts runFonts73 = new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi };
                FontSize fontSize69 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "18" };

                styleRunProperties6.Append(runFonts73);
                styleRunProperties6.Append(fontSize69);
                styleRunProperties6.Append(fontSizeComplexScript70);

                style12.Append(styleName12);
                style12.Append(basedOn8);
                style12.Append(linkedStyle8);
                style12.Append(uIPriority11);
                style12.Append(semiHidden7);
                style12.Append(rsid121);
                style12.Append(styleRunProperties6);

                Style style13 = new Style() { Type = StyleValues.Table, StyleId = "ab" };
                StyleName styleName13 = new StyleName() { Val = "Table Grid" };
                BasedOn basedOn9 = new BasedOn() { Val = "a1" };
                UIPriority uIPriority12 = new UIPriority() { Val = 59 };
                Rsid rsid122 = new Rsid() { Val = "00065B4E" };

                StyleTableProperties styleTableProperties2 = new StyleTableProperties();

                TableBorders tableBorders1 = new TableBorders();
                TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

                tableBorders1.Append(topBorder2);
                tableBorders1.Append(leftBorder2);
                tableBorders1.Append(bottomBorder2);
                tableBorders1.Append(rightBorder2);
                tableBorders1.Append(insideHorizontalBorder1);
                tableBorders1.Append(insideVerticalBorder1);

                styleTableProperties2.Append(tableBorders1);

                style13.Append(styleName13);
                style13.Append(basedOn9);
                style13.Append(uIPriority12);
                style13.Append(rsid122);
                style13.Append(styleTableProperties2);

                Style style14 = new Style() { Type = StyleValues.Paragraph, StyleId = "ac" };
                StyleName styleName14 = new StyleName() { Val = "List Paragraph" };
                BasedOn basedOn10 = new BasedOn() { Val = "a" };
                UIPriority uIPriority13 = new UIPriority() { Val = 34 };
                PrimaryStyle primaryStyle2 = new PrimaryStyle();
                Rsid rsid123 = new Rsid() { Val = "00104B80" };

                StyleParagraphProperties styleParagraphProperties5 = new StyleParagraphProperties();
                Indentation indentation10 = new Indentation() { Start = "480", StartCharacters = 200 };

                styleParagraphProperties5.Append(indentation10);

                style14.Append(styleName14);
                style14.Append(basedOn10);
                style14.Append(uIPriority13);
                style14.Append(primaryStyle2);
                style14.Append(rsid123);
                style14.Append(styleParagraphProperties5);

                Style style15 = new Style() { Type = StyleValues.Character, StyleId = "ad" };
                StyleName styleName15 = new StyleName() { Val = "annotation reference" };
                BasedOn basedOn11 = new BasedOn() { Val = "a0" };
                UIPriority uIPriority14 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden8 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed8 = new UnhideWhenUsed();
                Rsid rsid124 = new Rsid() { Val = "00691BB8" };

                StyleRunProperties styleRunProperties7 = new StyleRunProperties();
                FontSize fontSize70 = new FontSize() { Val = "18" };
                FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "18" };

                styleRunProperties7.Append(fontSize70);
                styleRunProperties7.Append(fontSizeComplexScript71);

                style15.Append(styleName15);
                style15.Append(basedOn11);
                style15.Append(uIPriority14);
                style15.Append(semiHidden8);
                style15.Append(unhideWhenUsed8);
                style15.Append(rsid124);
                style15.Append(styleRunProperties7);

                Style style16 = new Style() { Type = StyleValues.Paragraph, StyleId = "ae" };
                StyleName styleName16 = new StyleName() { Val = "annotation text" };
                BasedOn basedOn12 = new BasedOn() { Val = "a" };
                LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "af" };
                UIPriority uIPriority15 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden9 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed9 = new UnhideWhenUsed();
                Rsid rsid125 = new Rsid() { Val = "00691BB8" };

                style16.Append(styleName16);
                style16.Append(basedOn12);
                style16.Append(linkedStyle9);
                style16.Append(uIPriority15);
                style16.Append(semiHidden9);
                style16.Append(unhideWhenUsed9);
                style16.Append(rsid125);

                Style style17 = new Style() { Type = StyleValues.Character, StyleId = "af", CustomStyle = true };
                StyleName styleName17 = new StyleName() { Val = "註解文字 字元" };
                BasedOn basedOn13 = new BasedOn() { Val = "a0" };
                LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "ae" };
                UIPriority uIPriority16 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden10 = new SemiHidden();
                Rsid rsid126 = new Rsid() { Val = "00691BB8" };

                style17.Append(styleName17);
                style17.Append(basedOn13);
                style17.Append(linkedStyle10);
                style17.Append(uIPriority16);
                style17.Append(semiHidden10);
                style17.Append(rsid126);

                Style style18 = new Style() { Type = StyleValues.Paragraph, StyleId = "af0" };
                StyleName styleName18 = new StyleName() { Val = "annotation subject" };
                BasedOn basedOn14 = new BasedOn() { Val = "ae" };
                NextParagraphStyle nextParagraphStyle2 = new NextParagraphStyle() { Val = "ae" };
                LinkedStyle linkedStyle11 = new LinkedStyle() { Val = "af1" };
                UIPriority uIPriority17 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden11 = new SemiHidden();
                UnhideWhenUsed unhideWhenUsed10 = new UnhideWhenUsed();
                Rsid rsid127 = new Rsid() { Val = "00691BB8" };

                StyleRunProperties styleRunProperties8 = new StyleRunProperties();
                Bold bold4 = new Bold();
                BoldComplexScript boldComplexScript1 = new BoldComplexScript();

                styleRunProperties8.Append(bold4);
                styleRunProperties8.Append(boldComplexScript1);

                style18.Append(styleName18);
                style18.Append(basedOn14);
                style18.Append(nextParagraphStyle2);
                style18.Append(linkedStyle11);
                style18.Append(uIPriority17);
                style18.Append(semiHidden11);
                style18.Append(unhideWhenUsed10);
                style18.Append(rsid127);
                style18.Append(styleRunProperties8);

                Style style19 = new Style() { Type = StyleValues.Character, StyleId = "af1", CustomStyle = true };
                StyleName styleName19 = new StyleName() { Val = "註解主旨 字元" };
                BasedOn basedOn15 = new BasedOn() { Val = "af" };
                LinkedStyle linkedStyle12 = new LinkedStyle() { Val = "af0" };
                UIPriority uIPriority18 = new UIPriority() { Val = 99 };
                SemiHidden semiHidden12 = new SemiHidden();
                Rsid rsid128 = new Rsid() { Val = "00691BB8" };

                StyleRunProperties styleRunProperties9 = new StyleRunProperties();
                Bold bold5 = new Bold();
                BoldComplexScript boldComplexScript2 = new BoldComplexScript();

                styleRunProperties9.Append(bold5);
                styleRunProperties9.Append(boldComplexScript2);

                style19.Append(styleName19);
                style19.Append(basedOn15);
                style19.Append(linkedStyle12);
                style19.Append(uIPriority18);
                style19.Append(semiHidden12);
                style19.Append(rsid128);
                style19.Append(styleRunProperties9);

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
                styles1.Append(style9);
                styles1.Append(style10);
                styles1.Append(style11);
                styles1.Append(style12);
                styles1.Append(style13);
                styles1.Append(style14);
                styles1.Append(style15);
                styles1.Append(style16);
                styles1.Append(style17);
                styles1.Append(style18);
                styles1.Append(style19);

                styleDefinitionsPart1.Styles = styles1;
            }

            // Generates content of headerPart1.
            private void GenerateHeaderPart1Content(HeaderPart headerPart1)
            {
                Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
                header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                header1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
                header1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
                header1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
                header1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
                header1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
                header1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
                header1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
                header1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
                header1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
                header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                header1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
                header1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
                header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                header1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
                header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                header1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                header1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                header1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                header1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                header1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

                Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00397F67", RsidParagraphAddition = "00397F67", RsidParagraphProperties = "00D66A62", RsidRunAdditionDefault = "00397F67", ParagraphId = "6422DDEC", TextId = "043D62DE" };

                ParagraphProperties paragraphProperties22 = new ParagraphProperties();
                Justification justification7 = new Justification() { Val = JustificationValues.Right };

                ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
                RunFonts runFonts74 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
                Color color25 = new Color() { Val = "0000FF" };
                FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "28" };

                paragraphMarkRunProperties22.Append(runFonts74);
                paragraphMarkRunProperties22.Append(color25);
                paragraphMarkRunProperties22.Append(fontSizeComplexScript72);

                paragraphProperties22.Append(justification7);
                paragraphProperties22.Append(paragraphMarkRunProperties22);

                Run run50 = new Run() { RsidRunProperties = "009E00FF" };

                RunProperties runProperties50 = new RunProperties();
                RunFonts runFonts75 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體", EastAsia = "標楷體" };

                runProperties50.Append(runFonts75);
                Text text47 = new Text();
                text47.Text = "資安等級：內部";

                run50.Append(runProperties50);
                run50.Append(text47);

                paragraph22.Append(paragraphProperties22);
                paragraph22.Append(run50);

                Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "00397F67", RsidParagraphProperties = "00397F67", RsidRunAdditionDefault = "00397F67", ParagraphId = "2662779B", TextId = "77777777" };

                ParagraphProperties paragraphProperties23 = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "a5" };
                Justification justification8 = new Justification() { Val = JustificationValues.Right };

                paragraphProperties23.Append(paragraphStyleId5);
                paragraphProperties23.Append(justification8);

                paragraph23.Append(paragraphProperties23);

                header1.Append(paragraph22);
                header1.Append(paragraph23);

                headerPart1.Header = header1;
            }

            // Generates content of numberingDefinitionsPart1.
            private void GenerateNumberingDefinitionsPart1Content(NumberingDefinitionsPart numberingDefinitionsPart1)
            {
                Numbering numbering1 = new Numbering() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
                numbering1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                numbering1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
                numbering1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
                numbering1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
                numbering1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
                numbering1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
                numbering1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
                numbering1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
                numbering1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
                numbering1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
                numbering1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                numbering1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
                numbering1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
                numbering1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                numbering1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
                numbering1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                numbering1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                numbering1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                numbering1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                numbering1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                numbering1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                numbering1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                numbering1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                numbering1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                numbering1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                numbering1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                numbering1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                numbering1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                numbering1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                numbering1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                numbering1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                numbering1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                numbering1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

                AbstractNum abstractNum1 = new AbstractNum() { AbstractNumberId = 0 };
                abstractNum1.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
                Nsid nsid1 = new Nsid() { Val = "6E7A073A" };
                MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
                TemplateCode templateCode1 = new TemplateCode() { Val = "F020953E" };

                Level level1 = new Level() { LevelIndex = 0, TemplateCode = "9ABA3CA0" };
                StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText1 = new LevelText() { Val = "§" };
                LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
                Indentation indentation11 = new Indentation() { Start = "960", Hanging = "480" };

                previousParagraphProperties1.Append(indentation11);

                NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
                RunFonts runFonts76 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties1.Append(runFonts76);

                level1.Append(startNumberingValue1);
                level1.Append(numberingFormat1);
                level1.Append(levelText1);
                level1.Append(levelJustification1);
                level1.Append(previousParagraphProperties1);
                level1.Append(numberingSymbolRunProperties1);

                Level level2 = new Level() { LevelIndex = 1, TemplateCode = "04090003" };
                StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText2 = new LevelText() { Val = "n" };
                LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
                Indentation indentation12 = new Indentation() { Start = "1440", Hanging = "480" };

                previousParagraphProperties2.Append(indentation12);

                NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
                RunFonts runFonts77 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties2.Append(runFonts77);

                level2.Append(startNumberingValue2);
                level2.Append(numberingFormat2);
                level2.Append(levelText2);
                level2.Append(levelJustification2);
                level2.Append(previousParagraphProperties2);
                level2.Append(numberingSymbolRunProperties2);

                Level level3 = new Level() { LevelIndex = 2, TemplateCode = "04090005" };
                StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText3 = new LevelText() { Val = "u" };
                LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
                Indentation indentation13 = new Indentation() { Start = "1920", Hanging = "480" };

                previousParagraphProperties3.Append(indentation13);

                NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
                RunFonts runFonts78 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties3.Append(runFonts78);

                level3.Append(startNumberingValue3);
                level3.Append(numberingFormat3);
                level3.Append(levelText3);
                level3.Append(levelJustification3);
                level3.Append(previousParagraphProperties3);
                level3.Append(numberingSymbolRunProperties3);

                Level level4 = new Level() { LevelIndex = 3, TemplateCode = "04090001" };
                StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText4 = new LevelText() { Val = "l" };
                LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();
                Indentation indentation14 = new Indentation() { Start = "2400", Hanging = "480" };

                previousParagraphProperties4.Append(indentation14);

                NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
                RunFonts runFonts79 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties4.Append(runFonts79);

                level4.Append(startNumberingValue4);
                level4.Append(numberingFormat4);
                level4.Append(levelText4);
                level4.Append(levelJustification4);
                level4.Append(previousParagraphProperties4);
                level4.Append(numberingSymbolRunProperties4);

                Level level5 = new Level() { LevelIndex = 4, TemplateCode = "04090003", Tentative = true };
                StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText5 = new LevelText() { Val = "n" };
                LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();
                Indentation indentation15 = new Indentation() { Start = "2880", Hanging = "480" };

                previousParagraphProperties5.Append(indentation15);

                NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
                RunFonts runFonts80 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties5.Append(runFonts80);

                level5.Append(startNumberingValue5);
                level5.Append(numberingFormat5);
                level5.Append(levelText5);
                level5.Append(levelJustification5);
                level5.Append(previousParagraphProperties5);
                level5.Append(numberingSymbolRunProperties5);

                Level level6 = new Level() { LevelIndex = 5, TemplateCode = "04090005", Tentative = true };
                StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText6 = new LevelText() { Val = "u" };
                LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();
                Indentation indentation16 = new Indentation() { Start = "3360", Hanging = "480" };

                previousParagraphProperties6.Append(indentation16);

                NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
                RunFonts runFonts81 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties6.Append(runFonts81);

                level6.Append(startNumberingValue6);
                level6.Append(numberingFormat6);
                level6.Append(levelText6);
                level6.Append(levelJustification6);
                level6.Append(previousParagraphProperties6);
                level6.Append(numberingSymbolRunProperties6);

                Level level7 = new Level() { LevelIndex = 6, TemplateCode = "04090001", Tentative = true };
                StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText7 = new LevelText() { Val = "l" };
                LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();
                Indentation indentation17 = new Indentation() { Start = "3840", Hanging = "480" };

                previousParagraphProperties7.Append(indentation17);

                NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
                RunFonts runFonts82 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties7.Append(runFonts82);

                level7.Append(startNumberingValue7);
                level7.Append(numberingFormat7);
                level7.Append(levelText7);
                level7.Append(levelJustification7);
                level7.Append(previousParagraphProperties7);
                level7.Append(numberingSymbolRunProperties7);

                Level level8 = new Level() { LevelIndex = 7, TemplateCode = "04090003", Tentative = true };
                StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText8 = new LevelText() { Val = "n" };
                LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();
                Indentation indentation18 = new Indentation() { Start = "4320", Hanging = "480" };

                previousParagraphProperties8.Append(indentation18);

                NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
                RunFonts runFonts83 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties8.Append(runFonts83);

                level8.Append(startNumberingValue8);
                level8.Append(numberingFormat8);
                level8.Append(levelText8);
                level8.Append(levelJustification8);
                level8.Append(previousParagraphProperties8);
                level8.Append(numberingSymbolRunProperties8);

                Level level9 = new Level() { LevelIndex = 8, TemplateCode = "04090005", Tentative = true };
                StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
                NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
                LevelText levelText9 = new LevelText() { Val = "u" };
                LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Left };

                PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();
                Indentation indentation19 = new Indentation() { Start = "4800", Hanging = "480" };

                previousParagraphProperties9.Append(indentation19);

                NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
                RunFonts runFonts84 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

                numberingSymbolRunProperties9.Append(runFonts84);

                level9.Append(startNumberingValue9);
                level9.Append(numberingFormat9);
                level9.Append(levelText9);
                level9.Append(levelJustification9);
                level9.Append(previousParagraphProperties9);
                level9.Append(numberingSymbolRunProperties9);

                abstractNum1.Append(nsid1);
                abstractNum1.Append(multiLevelType1);
                abstractNum1.Append(templateCode1);
                abstractNum1.Append(level1);
                abstractNum1.Append(level2);
                abstractNum1.Append(level3);
                abstractNum1.Append(level4);
                abstractNum1.Append(level5);
                abstractNum1.Append(level6);
                abstractNum1.Append(level7);
                abstractNum1.Append(level8);
                abstractNum1.Append(level9);

                NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
                numberingInstance1.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "162819319"));
                AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 0 };

                numberingInstance1.Append(abstractNumId1);

                numbering1.Append(abstractNum1);
                numbering1.Append(numberingInstance1);

                numberingDefinitionsPart1.Numbering = numbering1;
            }

            // Generates content of endnotesPart1.
            private void GenerateEndnotesPart1Content(EndnotesPart endnotesPart1)
            {
                Endnotes endnotes1 = new Endnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
                endnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                endnotes1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
                endnotes1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
                endnotes1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
                endnotes1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
                endnotes1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
                endnotes1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
                endnotes1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
                endnotes1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
                endnotes1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
                endnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                endnotes1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
                endnotes1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
                endnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                endnotes1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
                endnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                endnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                endnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                endnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                endnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                endnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                endnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                endnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                endnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                endnotes1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                endnotes1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                endnotes1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                endnotes1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                endnotes1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                endnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                endnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                endnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                endnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

                Endnote endnote1 = new Endnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

                Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "00CC5221", RsidParagraphProperties = "00A66095", RsidRunAdditionDefault = "00CC5221", ParagraphId = "5DD76A5D", TextId = "77777777" };

                Run run51 = new Run();
                SeparatorMark separatorMark1 = new SeparatorMark();

                run51.Append(separatorMark1);

                paragraph24.Append(run51);

                endnote1.Append(paragraph24);

                Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

                Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "00CC5221", RsidParagraphProperties = "00A66095", RsidRunAdditionDefault = "00CC5221", ParagraphId = "3E41C675", TextId = "77777777" };

                Run run52 = new Run();
                ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

                run52.Append(continuationSeparatorMark1);

                paragraph25.Append(run52);

                endnote2.Append(paragraph25);

                endnotes1.Append(endnote1);
                endnotes1.Append(endnote2);

                endnotesPart1.Endnotes = endnotes1;
            }

            // Generates content of customXmlPart4.
            private void GenerateCustomXmlPart4Content(CustomXmlPart customXmlPart4)
            {
                System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart4.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><b:Sources xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" SelectedStyle=\"\\APASixthEditionOfficeOnline.xsl\" StyleName=\"APA\" Version=\"6\"></b:Sources>");
                writer.Flush();
                writer.Close();
            }

            // Generates content of customXmlPropertiesPart4.
            private void GenerateCustomXmlPropertiesPart4Content(CustomXmlPropertiesPart customXmlPropertiesPart4)
            {
                Ds.DataStoreItem dataStoreItem4 = new Ds.DataStoreItem() { ItemId = "{6433FADD-6D96-4DA6-9BA3-751AC4889DE9}" };
                dataStoreItem4.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

                Ds.SchemaReferences schemaReferences4 = new Ds.SchemaReferences();
                Ds.SchemaReference schemaReference15 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography" };

                schemaReferences4.Append(schemaReference15);

                dataStoreItem4.Append(schemaReferences4);

                customXmlPropertiesPart4.DataStoreItem = dataStoreItem4;
            }

            // Generates content of footnotesPart1.
            private void GenerateFootnotesPart1Content(FootnotesPart footnotesPart1)
            {
                Footnotes footnotes1 = new Footnotes() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
                footnotes1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
                footnotes1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
                footnotes1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
                footnotes1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
                footnotes1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
                footnotes1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
                footnotes1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
                footnotes1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
                footnotes1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
                footnotes1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
                footnotes1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
                footnotes1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
                footnotes1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
                footnotes1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
                footnotes1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
                footnotes1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                footnotes1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
                footnotes1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
                footnotes1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
                footnotes1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
                footnotes1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
                footnotes1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                footnotes1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                footnotes1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
                footnotes1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
                footnotes1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
                footnotes1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
                footnotes1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
                footnotes1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
                footnotes1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
                footnotes1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
                footnotes1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
                footnotes1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

                Footnote footnote1 = new Footnote() { Type = FootnoteEndnoteValues.Separator, Id = -1 };

                Paragraph paragraph26 = new Paragraph() { RsidParagraphAddition = "00CC5221", RsidParagraphProperties = "00A66095", RsidRunAdditionDefault = "00CC5221", ParagraphId = "7A6F3891", TextId = "77777777" };

                Run run53 = new Run();
                SeparatorMark separatorMark2 = new SeparatorMark();

                run53.Append(separatorMark2);

                paragraph26.Append(run53);

                footnote1.Append(paragraph26);

                Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

                Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "00CC5221", RsidParagraphProperties = "00A66095", RsidRunAdditionDefault = "00CC5221", ParagraphId = "20C8D974", TextId = "77777777" };

                Run run54 = new Run();
                ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

                run54.Append(continuationSeparatorMark2);

                paragraph27.Append(run54);

                footnote2.Append(paragraph27);

                footnotes1.Append(footnote1);
                footnotes1.Append(footnote2);

                footnotesPart1.Footnotes = footnotes1;
            }

            // Generates content of customFilePropertiesPart1.
            private void GenerateCustomFilePropertiesPart1Content(CustomFilePropertiesPart customFilePropertiesPart1)
            {
                Op.Properties properties2 = new Op.Properties();
                properties2.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

                Op.CustomDocumentProperty customDocumentProperty1 = new Op.CustomDocumentProperty() { FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", PropertyId = 2, Name = "ContentTypeId" };
                Vt.VTLPWSTR vTLPWSTR1 = new Vt.VTLPWSTR();
                vTLPWSTR1.Text = "0x010100D465D9393F8B7B429C8C9CE9DCEB2AA6";

                customDocumentProperty1.Append(vTLPWSTR1);

                properties2.Append(customDocumentProperty1);

                customFilePropertiesPart1.Properties = properties2;
            }

            private void SetPackageProperties(OpenXmlPackage document)
            {
                document.PackageProperties.Creator = "范添喜(Frank Fan)";
                document.PackageProperties.Revision = "30";
                document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2023-05-08T06:31:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
                document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2023-07-24T01:32:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
                document.PackageProperties.LastModifiedBy = "聖翔 王";
                document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2023-02-06T06:22:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            }


        }
    }
