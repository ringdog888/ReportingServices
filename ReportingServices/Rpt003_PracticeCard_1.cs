using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using System.Data;

namespace Rpt003_PracticeCard1
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
                words1.Text = "31";
                Ap.Characters characters1 = new Ap.Characters();
                characters1.Text = "179";
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
                charactersWithSpaces1.Text = "209";
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

                Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties1 = new ParagraphProperties();
                FrameProperties frameProperties1 = new FrameProperties() { Width = "360", Height = (UInt32Value)1920U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "4921", Y = "1486", HeightType = HeightRuleValues.Exact };
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
                paragraphProperties1.Append(autoSpaceDE1);
                paragraphProperties1.Append(autoSpaceDN1);
                paragraphProperties1.Append(adjustRightIndent1);
                paragraphProperties1.Append(paragraphMarkRunProperties1);
                ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run1 = new Run();

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
                text1.Text = "稽";

                run1.Append(runProperties1);
                run1.Append(text1);
                ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run2 = new Run();

                RunProperties runProperties2 = new RunProperties();
                RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color2 = new Color() { Val = "000000" };
                Kern kern3 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize2 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "24" };

                runProperties2.Append(runFonts3);
                runProperties2.Append(color2);
                runProperties2.Append(kern3);
                runProperties2.Append(fontSize2);
                runProperties2.Append(fontSizeComplexScript3);
                Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text2.Text = "　　　核";

                run2.Append(runProperties2);
                run2.Append(text2);

                paragraph1.Append(paragraphProperties1);
                paragraph1.Append(proofError1);
                paragraph1.Append(run1);
                paragraph1.Append(proofError2);
                paragraph1.Append(run2);

                Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties2 = new ParagraphProperties();
                FrameProperties frameProperties2 = new FrameProperties() { Width = "360", Height = (UInt32Value)1920U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "7561", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE2 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN2 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent2 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
                RunFonts runFonts4 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern4 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties2.Append(runFonts4);
                paragraphMarkRunProperties2.Append(kern4);
                paragraphMarkRunProperties2.Append(fontSizeComplexScript4);

                paragraphProperties2.Append(frameProperties2);
                paragraphProperties2.Append(autoSpaceDE2);
                paragraphProperties2.Append(autoSpaceDN2);
                paragraphProperties2.Append(adjustRightIndent2);
                paragraphProperties2.Append(paragraphMarkRunProperties2);

                Run run3 = new Run();

                RunProperties runProperties3 = new RunProperties();
                RunFonts runFonts5 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color3 = new Color() { Val = "000000" };
                Kern kern5 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize3 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "24" };

                runProperties3.Append(runFonts5);
                runProperties3.Append(color3);
                runProperties3.Append(kern5);
                runProperties3.Append(fontSize3);
                runProperties3.Append(fontSizeComplexScript5);
                Text text3 = new Text();
                text3.Text = "經　　　理";

                run3.Append(runProperties3);
                run3.Append(text3);

                paragraph2.Append(paragraphProperties2);
                paragraph2.Append(run3);

                Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties3 = new ParagraphProperties();
                FrameProperties frameProperties3 = new FrameProperties() { Width = "360", Height = (UInt32Value)1920U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "8881", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE3 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN3 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent3 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
                RunFonts runFonts6 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern6 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties3.Append(runFonts6);
                paragraphMarkRunProperties3.Append(kern6);
                paragraphMarkRunProperties3.Append(fontSizeComplexScript6);

                paragraphProperties3.Append(frameProperties3);
                paragraphProperties3.Append(autoSpaceDE3);
                paragraphProperties3.Append(autoSpaceDN3);
                paragraphProperties3.Append(adjustRightIndent3);
                paragraphProperties3.Append(paragraphMarkRunProperties3);

                Run run4 = new Run();

                RunProperties runProperties4 = new RunProperties();
                RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color4 = new Color() { Val = "000000" };
                Kern kern7 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize4 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "24" };

                runProperties4.Append(runFonts7);
                runProperties4.Append(color4);
                runProperties4.Append(kern7);
                runProperties4.Append(fontSize4);
                runProperties4.Append(fontSizeComplexScript7);
                Text text4 = new Text();
                text4.Text = "處　　　長";

                run4.Append(runProperties4);
                run4.Append(text4);

                paragraph3.Append(paragraphProperties3);
                paragraph3.Append(run4);

                Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();
                FrameProperties frameProperties4 = new FrameProperties() { Width = "360", Height = (UInt32Value)1920U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "10201", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE4 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN4 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent4 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                RunFonts runFonts8 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern8 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties4.Append(runFonts8);
                paragraphMarkRunProperties4.Append(kern8);
                paragraphMarkRunProperties4.Append(fontSizeComplexScript8);

                paragraphProperties4.Append(frameProperties4);
                paragraphProperties4.Append(autoSpaceDE4);
                paragraphProperties4.Append(autoSpaceDN4);
                paragraphProperties4.Append(adjustRightIndent4);
                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run5 = new Run();

                RunProperties runProperties5 = new RunProperties();
                RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color5 = new Color() { Val = "000000" };
                Kern kern9 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize5 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "24" };

                runProperties5.Append(runFonts9);
                runProperties5.Append(color5);
                runProperties5.Append(kern9);
                runProperties5.Append(fontSize5);
                runProperties5.Append(fontSizeComplexScript9);
                Text text5 = new Text();
                text5.Text = "總";

                run5.Append(runProperties5);
                run5.Append(text5);

                paragraph4.Append(paragraphProperties4);
                paragraph4.Append(run5);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties5 = new ParagraphProperties();
                FrameProperties frameProperties5 = new FrameProperties() { Width = "360", Height = (UInt32Value)1920U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "10201", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE5 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN5 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent5 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
                RunFonts runFonts10 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern10 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties5.Append(runFonts10);
                paragraphMarkRunProperties5.Append(kern10);
                paragraphMarkRunProperties5.Append(fontSizeComplexScript10);

                paragraphProperties5.Append(frameProperties5);
                paragraphProperties5.Append(autoSpaceDE5);
                paragraphProperties5.Append(autoSpaceDN5);
                paragraphProperties5.Append(adjustRightIndent5);
                paragraphProperties5.Append(paragraphMarkRunProperties5);

                Run run6 = new Run();

                RunProperties runProperties6 = new RunProperties();
                RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color6 = new Color() { Val = "000000" };
                Kern kern11 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize6 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "24" };

                runProperties6.Append(runFonts11);
                runProperties6.Append(color6);
                runProperties6.Append(kern11);
                runProperties6.Append(fontSize6);
                runProperties6.Append(fontSizeComplexScript11);
                Text text6 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text6.Text = "　";

                run6.Append(runProperties6);
                run6.Append(text6);
                ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run7 = new Run();

                RunProperties runProperties7 = new RunProperties();
                RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color7 = new Color() { Val = "000000" };
                Kern kern12 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize7 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "24" };

                runProperties7.Append(runFonts12);
                runProperties7.Append(color7);
                runProperties7.Append(kern12);
                runProperties7.Append(fontSize7);
                runProperties7.Append(fontSizeComplexScript12);
                Text text7 = new Text();
                text7.Text = "稽";

                run7.Append(runProperties7);
                run7.Append(text7);
                ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run8 = new Run();

                RunProperties runProperties8 = new RunProperties();
                RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color8 = new Color() { Val = "000000" };
                Kern kern13 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize8 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "24" };

                runProperties8.Append(runFonts13);
                runProperties8.Append(color8);
                runProperties8.Append(kern13);
                runProperties8.Append(fontSize8);
                runProperties8.Append(fontSizeComplexScript13);
                Text text8 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text8.Text = "　核";

                run8.Append(runProperties8);
                run8.Append(text8);

                paragraph5.Append(paragraphProperties5);
                paragraph5.Append(run6);
                paragraph5.Append(proofError3);
                paragraph5.Append(run7);
                paragraph5.Append(proofError4);
                paragraph5.Append(run8);

                Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "005416E1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties6 = new ParagraphProperties();
                FrameProperties frameProperties6 = new FrameProperties() { Width = "240", Height = (UInt32Value)1440U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "3001", Y = "2626", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE6 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN6 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent6 = new AdjustRightIndent() { Val = false };
                Justification justification1 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
                RunFonts runFonts14 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern14 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties6.Append(runFonts14);
                paragraphMarkRunProperties6.Append(kern14);
                paragraphMarkRunProperties6.Append(fontSizeComplexScript14);

                paragraphProperties6.Append(frameProperties6);
                paragraphProperties6.Append(autoSpaceDE6);
                paragraphProperties6.Append(autoSpaceDN6);
                paragraphProperties6.Append(adjustRightIndent6);
                paragraphProperties6.Append(justification1);
                paragraphProperties6.Append(paragraphMarkRunProperties6);

                Run run9 = new Run();

                RunProperties runProperties9 = new RunProperties();
                RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color9 = new Color() { Val = "000000" };
                Kern kern15 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize9 = new FontSize() { Val = "19" };
                FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "24" };

                runProperties9.Append(runFonts15);
                runProperties9.Append(color9);
                runProperties9.Append(kern15);
                runProperties9.Append(fontSize9);
                runProperties9.Append(fontSizeComplexScript15);
                Text text9 = new Text();
                text9.Text = "實習查核人";

                run9.Append(runProperties9);
                run9.Append(text9);

                Run run10 = new Run();

                RunProperties runProperties10 = new RunProperties();
                RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color10 = new Color() { Val = "000000" };
                Kern kern16 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize10 = new FontSize() { Val = "17" };
                FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "24" };

                runProperties10.Append(runFonts16);
                runProperties10.Append(color10);
                runProperties10.Append(kern16);
                runProperties10.Append(fontSize10);
                runProperties10.Append(fontSizeComplexScript16);
                Text text10 = new Text();
                text10.Text = "員";

                run10.Append(runProperties10);
                run10.Append(text10);

                paragraph6.Append(paragraphProperties6);
                paragraph6.Append(run9);
                paragraph6.Append(run10);

                Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "005416E1", RsidRunAdditionDefault = "005416E1" };

                ParagraphProperties paragraphProperties7 = new ParagraphProperties();
                FrameProperties frameProperties7 = new FrameProperties() { Width = "240", Height = (UInt32Value)1171U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "3001", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE7 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN7 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent7 = new AdjustRightIndent() { Val = false };
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "220", LineRule = LineSpacingRuleValues.Exact };
                Justification justification2 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
                RunFonts runFonts17 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern17 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties7.Append(runFonts17);
                paragraphMarkRunProperties7.Append(kern17);
                paragraphMarkRunProperties7.Append(fontSizeComplexScript17);

                paragraphProperties7.Append(frameProperties7);
                paragraphProperties7.Append(autoSpaceDE7);
                paragraphProperties7.Append(autoSpaceDN7);
                paragraphProperties7.Append(adjustRightIndent7);
                paragraphProperties7.Append(spacingBetweenLines1);
                paragraphProperties7.Append(justification2);
                paragraphProperties7.Append(paragraphMarkRunProperties7);

                Run run11 = new Run();

                RunProperties runProperties11 = new RunProperties();
                RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color11 = new Color() { Val = "000000" };
                Kern kern18 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize11 = new FontSize() { Val = "19" };
                FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "24" };

                runProperties11.Append(runFonts18);
                runProperties11.Append(color11);
                runProperties11.Append(kern18);
                runProperties11.Append(fontSize11);
                runProperties11.Append(fontSizeComplexScript18);
                Text text11 = new Text();
                text11.Text = dt.Rows[0]["Issue"].ToString();

                run11.Append(runProperties11);
                run11.Append(text11);

                paragraph7.Append(paragraphProperties7);
                paragraph7.Append(run11);

                Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "005416E1", RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "005416E1", RsidRunAdditionDefault = "005416E1" };

                ParagraphProperties paragraphProperties8 = new ParagraphProperties();
                FrameProperties frameProperties8 = new FrameProperties() { Width = "240", Height = (UInt32Value)720U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "2641", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE8 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN8 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent8 = new AdjustRightIndent() { Val = false };
                SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "220", LineRule = LineSpacingRuleValues.Exact };
                Justification justification3 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
                RunFonts runFonts19 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern19 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties8.Append(runFonts19);
                paragraphMarkRunProperties8.Append(kern19);
                paragraphMarkRunProperties8.Append(fontSizeComplexScript19);

                paragraphProperties8.Append(frameProperties8);
                paragraphProperties8.Append(autoSpaceDE8);
                paragraphProperties8.Append(autoSpaceDN8);
                paragraphProperties8.Append(adjustRightIndent8);
                paragraphProperties8.Append(spacingBetweenLines2);
                paragraphProperties8.Append(justification3);
                paragraphProperties8.Append(paragraphMarkRunProperties8);

                Run run12 = new Run() { RsidRunProperties = "005416E1" };

                RunProperties runProperties12 = new RunProperties();
                RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color12 = new Color() { Val = "000000" };
                Kern kern20 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize12 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "24" };

                runProperties12.Append(runFonts20);
                runProperties12.Append(color12);
                runProperties12.Append(kern20);
                runProperties12.Append(fontSize12);
                runProperties12.Append(fontSizeComplexScript20);
                Text text12 = new Text();
                text12.Text = dt.Rows[0]["ppl"].ToString();

                run12.Append(runProperties12);
                run12.Append(text12);

                paragraph8.Append(paragraphProperties8);
                paragraph8.Append(run12);

                Paragraph paragraph9 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties9 = new ParagraphProperties();
                FrameProperties frameProperties9 = new FrameProperties() { Width = "240", Height = (UInt32Value)240U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "2641", Y = "2206", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE9 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN9 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent9 = new AdjustRightIndent() { Val = false };
                Justification justification4 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
                RunFonts runFonts21 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern21 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties9.Append(runFonts21);
                paragraphMarkRunProperties9.Append(kern21);
                paragraphMarkRunProperties9.Append(fontSizeComplexScript21);

                paragraphProperties9.Append(frameProperties9);
                paragraphProperties9.Append(autoSpaceDE9);
                paragraphProperties9.Append(autoSpaceDN9);
                paragraphProperties9.Append(adjustRightIndent9);
                paragraphProperties9.Append(justification4);
                paragraphProperties9.Append(paragraphMarkRunProperties9);

                Run run13 = new Run();

                RunProperties runProperties13 = new RunProperties();
                RunFonts runFonts22 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color13 = new Color() { Val = "000000" };
                Kern kern22 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize13 = new FontSize() { Val = "19" };
                FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "24" };

                runProperties13.Append(runFonts22);
                runProperties13.Append(color13);
                runProperties13.Append(kern22);
                runProperties13.Append(fontSize13);
                runProperties13.Append(fontSizeComplexScript22);
                Text text13 = new Text();
                text13.Text = "等";

                run13.Append(runProperties13);
                run13.Append(text13);

                paragraph9.Append(paragraphProperties9);
                paragraph9.Append(run13);

                Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties10 = new ParagraphProperties();
                FrameProperties frameProperties10 = new FrameProperties() { Width = "240", Height = (UInt32Value)880U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "2641", Y = "2766", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE10 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN10 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent10 = new AdjustRightIndent() { Val = false };
                Justification justification5 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
                RunFonts runFonts23 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern23 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties10.Append(runFonts23);
                paragraphMarkRunProperties10.Append(kern23);
                paragraphMarkRunProperties10.Append(fontSizeComplexScript23);

                paragraphProperties10.Append(frameProperties10);
                paragraphProperties10.Append(autoSpaceDE10);
                paragraphProperties10.Append(autoSpaceDN10);
                paragraphProperties10.Append(adjustRightIndent10);
                paragraphProperties10.Append(justification5);
                paragraphProperties10.Append(paragraphMarkRunProperties10);

                Run run14 = new Run();

                RunProperties runProperties14 = new RunProperties();
                RunFonts runFonts24 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color14 = new Color() { Val = "000000" };
                Kern kern24 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize14 = new FontSize() { Val = "17" };
                FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "24" };

                runProperties14.Append(runFonts24);
                runProperties14.Append(color14);
                runProperties14.Append(kern24);
                runProperties14.Append(fontSize14);
                runProperties14.Append(fontSizeComplexScript24);
                Text text14 = new Text();
                text14.Text = "員，已完";

                run14.Append(runProperties14);
                run14.Append(text14);

                paragraph10.Append(paragraphProperties10);
                paragraph10.Append(run14);

                Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties11 = new ParagraphProperties();
                FrameProperties frameProperties11 = new FrameProperties() { Width = "240", Height = (UInt32Value)2160U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "2281", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE11 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN11 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent11 = new AdjustRightIndent() { Val = false };
                Justification justification6 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
                RunFonts runFonts25 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern25 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties11.Append(runFonts25);
                paragraphMarkRunProperties11.Append(kern25);
                paragraphMarkRunProperties11.Append(fontSizeComplexScript25);

                paragraphProperties11.Append(frameProperties11);
                paragraphProperties11.Append(autoSpaceDE11);
                paragraphProperties11.Append(autoSpaceDN11);
                paragraphProperties11.Append(adjustRightIndent11);
                paragraphProperties11.Append(justification6);
                paragraphProperties11.Append(paragraphMarkRunProperties11);

                Run run15 = new Run();

                RunProperties runProperties15 = new RunProperties();
                RunFonts runFonts26 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color15 = new Color() { Val = "000000" };
                Kern kern26 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize15 = new FontSize() { Val = "17" };
                FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "24" };

                runProperties15.Append(runFonts26);
                runProperties15.Append(color15);
                runProperties15.Append(kern26);
                runProperties15.Append(fontSize15);
                runProperties15.Append(fontSizeComplexScript26);
                Text text15 = new Text();
                text15.Text = "成查核實習，成績及格";

                run15.Append(runProperties15);
                run15.Append(text15);

                paragraph11.Append(paragraphProperties11);
                paragraph11.Append(run15);

                Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties12 = new ParagraphProperties();
                FrameProperties frameProperties12 = new FrameProperties() { Width = "240", Height = (UInt32Value)2160U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1921", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE12 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN12 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent12 = new AdjustRightIndent() { Val = false };
                Justification justification7 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
                RunFonts runFonts27 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern27 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties12.Append(runFonts27);
                paragraphMarkRunProperties12.Append(kern27);
                paragraphMarkRunProperties12.Append(fontSizeComplexScript27);

                paragraphProperties12.Append(frameProperties12);
                paragraphProperties12.Append(autoSpaceDE12);
                paragraphProperties12.Append(autoSpaceDN12);
                paragraphProperties12.Append(adjustRightIndent12);
                paragraphProperties12.Append(justification7);
                paragraphProperties12.Append(paragraphMarkRunProperties12);

                Run run16 = new Run();

                RunProperties runProperties16 = new RunProperties();
                RunFonts runFonts28 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color16 = new Color() { Val = "000000" };
                Kern kern28 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize16 = new FontSize() { Val = "17" };
                FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "24" };

                runProperties16.Append(runFonts28);
                runProperties16.Append(color16);
                runProperties16.Append(kern28);
                runProperties16.Append(fontSize16);
                runProperties16.Append(fontSizeComplexScript28);
                Text text16 = new Text();
                text16.Text = "擬請";

                run16.Append(runProperties16);
                run16.Append(text16);

                paragraph12.Append(paragraphProperties12);
                paragraph12.Append(run16);

                Paragraph paragraph13 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties13 = new ParagraphProperties();
                FrameProperties frameProperties13 = new FrameProperties() { Width = "240", Height = (UInt32Value)2160U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1561", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE13 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN13 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent13 = new AdjustRightIndent() { Val = false };
                Justification justification8 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
                RunFonts runFonts29 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern29 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties13.Append(runFonts29);
                paragraphMarkRunProperties13.Append(kern29);
                paragraphMarkRunProperties13.Append(fontSizeComplexScript29);

                paragraphProperties13.Append(frameProperties13);
                paragraphProperties13.Append(autoSpaceDE13);
                paragraphProperties13.Append(autoSpaceDN13);
                paragraphProperties13.Append(adjustRightIndent13);
                paragraphProperties13.Append(justification8);
                paragraphProperties13.Append(paragraphMarkRunProperties13);

                Run run17 = new Run();

                RunProperties runProperties17 = new RunProperties();
                RunFonts runFonts30 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color17 = new Color() { Val = "000000" };
                Kern kern30 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize17 = new FontSize() { Val = "17" };
                FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "24" };

                runProperties17.Append(runFonts30);
                runProperties17.Append(color17);
                runProperties17.Append(kern30);
                runProperties17.Append(fontSize17);
                runProperties17.Append(fontSizeComplexScript30);
                Text text17 = new Text();
                text17.Text = "總稽核簽發證明書送";

                run17.Append(runProperties17);
                run17.Append(text17);

                paragraph13.Append(paragraphProperties13);
                paragraph13.Append(run17);

                Paragraph paragraph14 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties14 = new ParagraphProperties();
                FrameProperties frameProperties14 = new FrameProperties() { Width = "240", Height = (UInt32Value)2160U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1201", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE14 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN14 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent14 = new AdjustRightIndent() { Val = false };
                Justification justification9 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
                RunFonts runFonts31 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern31 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties14.Append(runFonts31);
                paragraphMarkRunProperties14.Append(kern31);
                paragraphMarkRunProperties14.Append(fontSizeComplexScript31);

                paragraphProperties14.Append(frameProperties14);
                paragraphProperties14.Append(autoSpaceDE14);
                paragraphProperties14.Append(autoSpaceDN14);
                paragraphProperties14.Append(adjustRightIndent14);
                paragraphProperties14.Append(justification9);
                paragraphProperties14.Append(paragraphMarkRunProperties14);

                Run run18 = new Run();

                RunProperties runProperties18 = new RunProperties();
                RunFonts runFonts32 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color18 = new Color() { Val = "000000" };
                Kern kern32 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize18 = new FontSize() { Val = "17" };
                FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "24" };

                runProperties18.Append(runFonts32);
                runProperties18.Append(color18);
                runProperties18.Append(kern32);
                runProperties18.Append(fontSize18);
                runProperties18.Append(fontSizeComplexScript32);
                Text text18 = new Text();
                text18.Text = "人力資源處";

                run18.Append(runProperties18);
                run18.Append(text18);

                paragraph14.Append(paragraphProperties14);
                paragraph14.Append(run18);

                Paragraph paragraph15 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties15 = new ParagraphProperties();
                FrameProperties frameProperties15 = new FrameProperties() { Width = "240", Height = (UInt32Value)2160U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1201", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE15 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN15 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent15 = new AdjustRightIndent() { Val = false };
                Justification justification10 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
                RunFonts runFonts33 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern33 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties15.Append(runFonts33);
                paragraphMarkRunProperties15.Append(kern33);
                paragraphMarkRunProperties15.Append(fontSizeComplexScript33);

                paragraphProperties15.Append(frameProperties15);
                paragraphProperties15.Append(autoSpaceDE15);
                paragraphProperties15.Append(autoSpaceDN15);
                paragraphProperties15.Append(adjustRightIndent15);
                paragraphProperties15.Append(justification10);
                paragraphProperties15.Append(paragraphMarkRunProperties15);

                Run run19 = new Run();

                RunProperties runProperties19 = new RunProperties();
                RunFonts runFonts34 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color19 = new Color() { Val = "000000" };
                Kern kern34 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize19 = new FontSize() { Val = "21" };
                FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "24" };

                runProperties19.Append(runFonts34);
                runProperties19.Append(color19);
                runProperties19.Append(kern34);
                runProperties19.Append(fontSize19);
                runProperties19.Append(fontSizeComplexScript34);
                Text text19 = new Text();
                text19.Text = "。";

                run19.Append(runProperties19);
                run19.Append(text19);

                paragraph15.Append(paragraphProperties15);
                paragraph15.Append(run19);

                Paragraph paragraph16 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "005416E1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties16 = new ParagraphProperties();
                FrameProperties frameProperties16 = new FrameProperties() { Width = "240", Height = (UInt32Value)240U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "2641", Y = "2486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE16 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN16 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent16 = new AdjustRightIndent() { Val = false };
                Justification justification11 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
                RunFonts runFonts35 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern35 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties16.Append(runFonts35);
                paragraphMarkRunProperties16.Append(kern35);
                paragraphMarkRunProperties16.Append(fontSizeComplexScript35);

                paragraphProperties16.Append(frameProperties16);
                paragraphProperties16.Append(autoSpaceDE16);
                paragraphProperties16.Append(autoSpaceDN16);
                paragraphProperties16.Append(adjustRightIndent16);
                paragraphProperties16.Append(justification11);
                paragraphProperties16.Append(paragraphMarkRunProperties16);

                Run run20 = new Run();

                RunProperties runProperties20 = new RunProperties();
                RunFonts runFonts36 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Color color20 = new Color() { Val = "000000" };
                Kern kern36 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize20 = new FontSize() { Val = "19" };
                FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "24" };

                runProperties20.Append(runFonts36);
                runProperties20.Append(color20);
                runProperties20.Append(kern36);
                runProperties20.Append(fontSize20);
                runProperties20.Append(fontSizeComplexScript36);
                Text text20 = new Text();
                text20.Text = dt.Rows[0]["p_num"].ToString();

                run20.Append(runProperties20);
                run20.Append(text20);

                paragraph16.Append(paragraphProperties16);
                paragraph16.Append(run20);

                Paragraph paragraph17 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties17 = new ParagraphProperties();
                FrameProperties frameProperties17 = new FrameProperties() { Width = "360", Height = (UInt32Value)1920U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "6241", Y = "1486", HeightType = HeightRuleValues.Exact };
                AutoSpaceDE autoSpaceDE17 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN17 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent17 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
                RunFonts runFonts37 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern37 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties17.Append(runFonts37);
                paragraphMarkRunProperties17.Append(kern37);
                paragraphMarkRunProperties17.Append(fontSizeComplexScript37);

                paragraphProperties17.Append(frameProperties17);
                paragraphProperties17.Append(autoSpaceDE17);
                paragraphProperties17.Append(autoSpaceDN17);
                paragraphProperties17.Append(adjustRightIndent17);
                paragraphProperties17.Append(paragraphMarkRunProperties17);

                Run run21 = new Run();

                RunProperties runProperties21 = new RunProperties();
                RunFonts runFonts38 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color21 = new Color() { Val = "000000" };
                Kern kern38 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize21 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "24" };

                runProperties21.Append(runFonts38);
                runProperties21.Append(color21);
                runProperties21.Append(kern38);
                runProperties21.Append(fontSize21);
                runProperties21.Append(fontSizeComplexScript38);
                Text text21 = new Text();
                text21.Text = "副　　　理";

                run21.Append(runProperties21);
                run21.Append(text21);

                paragraph17.Append(paragraphProperties17);
                paragraph17.Append(run21);

                Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "002D5CA1", RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "007158A7", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties18 = new ParagraphProperties();
                FrameProperties frameProperties18 = new FrameProperties() { Width = "7471", Height = (UInt32Value)375U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1201", Y = "4786", HeightType = HeightRuleValues.Exact };

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
                AutoSpaceDE autoSpaceDE18 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN18 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent18 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
                RunFonts runFonts39 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern39 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties18.Append(runFonts39);
                paragraphMarkRunProperties18.Append(kern39);
                paragraphMarkRunProperties18.Append(fontSizeComplexScript39);

                paragraphProperties18.Append(frameProperties18);
                paragraphProperties18.Append(tabs1);
                paragraphProperties18.Append(autoSpaceDE18);
                paragraphProperties18.Append(autoSpaceDN18);
                paragraphProperties18.Append(adjustRightIndent18);
                paragraphProperties18.Append(paragraphMarkRunProperties18);

                Run run22 = new Run() { RsidRunProperties = "002D5CA1" };

                RunProperties runProperties22 = new RunProperties();
                RunFonts runFonts40 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color22 = new Color() { Val = "000000" };
                Kern kern40 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize22 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "24" };

                runProperties22.Append(runFonts40);
                runProperties22.Append(color22);
                runProperties22.Append(kern40);
                runProperties22.Append(fontSize22);
                runProperties22.Append(fontSizeComplexScript40);
                Text text22 = new Text();
                text22.Text = "人力資源處台照：";

                run22.Append(runProperties22);
                run22.Append(text22);

                paragraph18.Append(paragraphProperties18);
                paragraph18.Append(run22);

                Paragraph paragraph19 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "007158A7", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties19 = new ParagraphProperties();
                FrameProperties frameProperties19 = new FrameProperties() { Width = "2040", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "6121", Y = "10921", HeightType = HeightRuleValues.Exact };

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
                AutoSpaceDE autoSpaceDE19 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN19 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent19 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
                RunFonts runFonts41 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern41 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties19.Append(runFonts41);
                paragraphMarkRunProperties19.Append(kern41);
                paragraphMarkRunProperties19.Append(fontSizeComplexScript41);

                paragraphProperties19.Append(frameProperties19);
                paragraphProperties19.Append(tabs2);
                paragraphProperties19.Append(autoSpaceDE19);
                paragraphProperties19.Append(autoSpaceDN19);
                paragraphProperties19.Append(adjustRightIndent19);
                paragraphProperties19.Append(paragraphMarkRunProperties19);

                Run run23 = new Run();

                RunProperties runProperties23 = new RunProperties();
                RunFonts runFonts42 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color23 = new Color() { Val = "000000" };
                Kern kern42 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize23 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "24" };

                runProperties23.Append(runFonts42);
                runProperties23.Append(color23);
                runProperties23.Append(kern42);
                runProperties23.Append(fontSize23);
                runProperties23.Append(fontSizeComplexScript42);
                Text text23 = new Text();
                text23.Text = "董事會稽核處";

                run23.Append(runProperties23);
                run23.Append(text23);

                paragraph19.Append(paragraphProperties19);
                paragraph19.Append(run23);

                Paragraph paragraph20 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "007158A7", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties20 = new ParagraphProperties();
                FrameProperties frameProperties20 = new FrameProperties() { Width = "2720", Height = (UInt32Value)360U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "5761", Y = "11596", HeightType = HeightRuleValues.Exact };

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
                AutoSpaceDE autoSpaceDE20 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN20 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent20 = new AdjustRightIndent() { Val = false };
                Justification justification12 = new Justification() { Val = JustificationValues.Center };

                ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
                RunFonts runFonts43 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern43 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties20.Append(runFonts43);
                paragraphMarkRunProperties20.Append(kern43);
                paragraphMarkRunProperties20.Append(fontSizeComplexScript43);

                paragraphProperties20.Append(frameProperties20);
                paragraphProperties20.Append(tabs3);
                paragraphProperties20.Append(autoSpaceDE20);
                paragraphProperties20.Append(autoSpaceDN20);
                paragraphProperties20.Append(adjustRightIndent20);
                paragraphProperties20.Append(justification12);
                paragraphProperties20.Append(paragraphMarkRunProperties20);

                Run run24 = new Run();

                RunProperties runProperties24 = new RunProperties();
                RunFonts runFonts44 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Color color24 = new Color() { Val = "000000" };
                Kern kern44 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize24 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "24" };

                runProperties24.Append(runFonts44);
                runProperties24.Append(color24);
                runProperties24.Append(kern44);
                runProperties24.Append(fontSize24);
                runProperties24.Append(fontSizeComplexScript44);
                Text text24 = new Text();
                text24.Text = dt.Rows[0]["Year"].ToString();

                run24.Append(runProperties24);
                run24.Append(text24);

                Run run25 = new Run();

                RunProperties runProperties25 = new RunProperties();
                RunFonts runFonts45 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color25 = new Color() { Val = "000000" };
                Kern kern45 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize25 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "24" };

                runProperties25.Append(runFonts45);
                runProperties25.Append(color25);
                runProperties25.Append(kern45);
                runProperties25.Append(fontSize25);
                runProperties25.Append(fontSizeComplexScript45);
                Text text25 = new Text();
                text25.Text = "年";

                run25.Append(runProperties25);
                run25.Append(text25);

                Run run26 = new Run();

                RunProperties runProperties26 = new RunProperties();
                RunFonts runFonts46 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Color color26 = new Color() { Val = "000000" };
                Kern kern46 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize26 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "24" };

                runProperties26.Append(runFonts46);
                runProperties26.Append(color26);
                runProperties26.Append(kern46);
                runProperties26.Append(fontSize26);
                runProperties26.Append(fontSizeComplexScript46);
                Text text26 = new Text();
                text26.Text = dt.Rows[0]["Month"].ToString();

                run26.Append(runProperties26);
                run26.Append(text26);

                Run run27 = new Run();

                RunProperties runProperties27 = new RunProperties();
                RunFonts runFonts47 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color27 = new Color() { Val = "000000" };
                Kern kern47 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize27 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "24" };

                runProperties27.Append(runFonts47);
                runProperties27.Append(color27);
                runProperties27.Append(kern47);
                runProperties27.Append(fontSize27);
                runProperties27.Append(fontSizeComplexScript47);
                Text text27 = new Text();
                text27.Text = "月";

                run27.Append(runProperties27);
                run27.Append(text27);

                Run run28 = new Run();

                RunProperties runProperties28 = new RunProperties();
                RunFonts runFonts48 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Color color28 = new Color() { Val = "000000" };
                Kern kern48 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize28 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "24" };

                runProperties28.Append(runFonts48);
                runProperties28.Append(color28);
                runProperties28.Append(kern48);
                runProperties28.Append(fontSize28);
                runProperties28.Append(fontSizeComplexScript48);
                Text text28 = new Text();
                text28.Text = dt.Rows[0]["Day"].ToString();

                run28.Append(runProperties28);
                run28.Append(text28);

                Run run29 = new Run();

                RunProperties runProperties29 = new RunProperties();
                RunFonts runFonts49 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color29 = new Color() { Val = "000000" };
                Kern kern49 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize29 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "24" };

                runProperties29.Append(runFonts49);
                runProperties29.Append(color29);
                runProperties29.Append(kern49);
                runProperties29.Append(fontSize29);
                runProperties29.Append(fontSizeComplexScript49);
                Text text29 = new Text();
                text29.Text = "日";

                run29.Append(runProperties29);
                run29.Append(text29);

                paragraph20.Append(paragraphProperties20);
                paragraph20.Append(run24);
                paragraph20.Append(run25);
                paragraph20.Append(run26);
                paragraph20.Append(run27);
                paragraph20.Append(run28);
                paragraph20.Append(run29);

                Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "005416E1", RsidParagraphProperties = "007158A7", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties21 = new ParagraphProperties();
                FrameProperties frameProperties21 = new FrameProperties() { Width = "7936", Height = (UInt32Value)4036U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1876", Y = "6121", HeightType = HeightRuleValues.Exact };

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
                AutoSpaceDE autoSpaceDE21 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN21 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent21 = new AdjustRightIndent() { Val = false };
                SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
                RunFonts runFonts50 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color30 = new Color() { Val = "000000" };
                Kern kern50 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize30 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties21.Append(runFonts50);
                paragraphMarkRunProperties21.Append(color30);
                paragraphMarkRunProperties21.Append(kern50);
                paragraphMarkRunProperties21.Append(fontSize30);
                paragraphMarkRunProperties21.Append(fontSizeComplexScript50);

                paragraphProperties21.Append(frameProperties21);
                paragraphProperties21.Append(tabs4);
                paragraphProperties21.Append(autoSpaceDE21);
                paragraphProperties21.Append(autoSpaceDN21);
                paragraphProperties21.Append(adjustRightIndent21);
                paragraphProperties21.Append(spacingBetweenLines3);
                paragraphProperties21.Append(paragraphMarkRunProperties21);

                Run run30 = new Run() { RsidRunProperties = "002D5CA1" };

                RunProperties runProperties30 = new RunProperties();
                RunFonts runFonts51 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color31 = new Color() { Val = "000000" };
                Kern kern51 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize31 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "24" };

                runProperties30.Append(runFonts51);
                runProperties30.Append(color31);
                runProperties30.Append(kern51);
                runProperties30.Append(fontSize31);
                runProperties30.Append(fontSizeComplexScript51);
                Text text30 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                text30.Text = "　　貴處";

                run30.Append(runProperties30);
                run30.Append(text30);

                Run run31 = new Run() { RsidRunProperties = "005416E1", RsidRunAddition = "005416E1" };

                RunProperties runProperties31 = new RunProperties();
                RunFonts runFonts52 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color32 = new Color() { Val = "000000" };
                Kern kern52 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize32 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "24" };

                runProperties31.Append(runFonts52);
                runProperties31.Append(color32);
                runProperties31.Append(kern52);
                runProperties31.Append(fontSize32);
                runProperties31.Append(fontSizeComplexScript52);
                Text text31 = new Text();
                text31.Text = dt.Rows[0]["context"].ToString();

                run31.Append(runProperties31);
                run31.Append(text31);
                ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run32 = new Run() { RsidRunProperties = "005416E1", RsidRunAddition = "005416E1" };

                RunProperties runProperties32 = new RunProperties();
                RunFonts runFonts53 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color33 = new Color() { Val = "000000" };
                Kern kern53 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize33 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "24" };

                runProperties32.Append(runFonts53);
                runProperties32.Append(color33);
                runProperties32.Append(kern53);
                runProperties32.Append(fontSize33);
                runProperties32.Append(fontSizeComplexScript53);
                Text text32 = new Text();
                text32.Text = "政人訓字";

                run32.Append(runProperties32);
                run32.Append(text32);
                ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run33 = new Run() { RsidRunProperties = "005416E1", RsidRunAddition = "005416E1" };

                RunProperties runProperties33 = new RunProperties();
                RunFonts runFonts54 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color34 = new Color() { Val = "000000" };
                Kern kern54 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize34 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "24" };

                runProperties33.Append(runFonts54);
                runProperties33.Append(color34);
                runProperties33.Append(kern54);
                runProperties33.Append(fontSize34);
                runProperties33.Append(fontSizeComplexScript54);
                Text text33 = new Text();
                text33.Text = "第00246號";

                run33.Append(runProperties33);
                run33.Append(text33);

                Run run34 = new Run() { RsidRunAddition = "005416E1" };

                RunProperties runProperties34 = new RunProperties();
                RunFonts runFonts55 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color35 = new Color() { Val = "000000" };
                Kern kern55 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize35 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "24" };

                runProperties34.Append(runFonts55);
                runProperties34.Append(color35);
                runProperties34.Append(kern55);
                runProperties34.Append(fontSize35);
                runProperties34.Append(fontSizeComplexScript55);
                Text text34 = new Text();
                text34.Text = "函";

                run34.Append(runProperties34);
                run34.Append(text34);

                paragraph21.Append(paragraphProperties21);
                paragraph21.Append(run30);
                paragraph21.Append(run31);
                paragraph21.Append(proofError5);
                paragraph21.Append(run32);
                paragraph21.Append(proofError6);
                paragraph21.Append(run33);
                paragraph21.Append(run34);

                Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "002D5CA1", RsidParagraphAddition = "002D5CA1", RsidParagraphProperties = "007158A7", RsidRunAdditionDefault = "005416E1" };

                ParagraphProperties paragraphProperties22 = new ParagraphProperties();
                FrameProperties frameProperties22 = new FrameProperties() { Width = "7936", Height = (UInt32Value)4036U, Wrap = TextWrappingValues.Auto, HorizontalPosition = HorizontalAnchorValues.Page, VerticalPosition = VerticalAnchorValues.Page, X = "1876", Y = "6121", HeightType = HeightRuleValues.Exact };

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
                AutoSpaceDE autoSpaceDE22 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN22 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent22 = new AdjustRightIndent() { Val = false };
                SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "480", LineRule = LineSpacingRuleValues.Auto };

                ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
                RunFonts runFonts56 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern56 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties22.Append(runFonts56);
                paragraphMarkRunProperties22.Append(kern56);
                paragraphMarkRunProperties22.Append(fontSizeComplexScript56);

                paragraphProperties22.Append(frameProperties22);
                paragraphProperties22.Append(tabs5);
                paragraphProperties22.Append(autoSpaceDE22);
                paragraphProperties22.Append(autoSpaceDN22);
                paragraphProperties22.Append(adjustRightIndent22);
                paragraphProperties22.Append(spacingBetweenLines4);
                paragraphProperties22.Append(paragraphMarkRunProperties22);

                Run run35 = new Run();

                RunProperties runProperties35 = new RunProperties();
                RunFonts runFonts57 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color36 = new Color() { Val = "000000" };
                Kern kern57 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize36 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "24" };

                runProperties35.Append(runFonts57);
                runProperties35.Append(color36);
                runProperties35.Append(kern57);
                runProperties35.Append(fontSize36);
                runProperties35.Append(fontSizeComplexScript57);
                Text text35 = new Text();
                text35.Text = "派查核實習人員" + dt.Rows[0]["ppl"].ToString();

                run35.Append(runProperties35);
                run35.Append(text35);

                Run run36 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties36 = new RunProperties();
                RunFonts runFonts58 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color37 = new Color() { Val = "000000" };
                Kern kern58 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize37 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "24" };

                runProperties36.Append(runFonts58);
                runProperties36.Append(color37);
                runProperties36.Append(kern58);
                runProperties36.Append(fontSize37);
                runProperties36.Append(fontSizeComplexScript58);
                Text text36 = new Text();
                text36.Text = "等";

                run36.Append(runProperties36);
                run36.Append(text36);

                Run run37 = new Run() { RsidRunAddition = "007158A7" };

                RunProperties runProperties37 = new RunProperties();
                RunFonts runFonts59 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color38 = new Color() { Val = "000000" };
                Kern kern59 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize38 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "24" };

                runProperties37.Append(runFonts59);
                runProperties37.Append(color38);
                runProperties37.Append(kern59);
                runProperties37.Append(fontSize38);
                runProperties37.Append(fontSizeComplexScript59);
                Text text37 = new Text();
                text37.Text = dt.Rows[0]["p_num"].ToString();

                run37.Append(runProperties37);
                run37.Append(text37);

                Run run38 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties38 = new RunProperties();
                RunFonts runFonts60 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color39 = new Color() { Val = "000000" };
                Kern kern60 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize39 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "24" };

                runProperties38.Append(runFonts60);
                runProperties38.Append(color39);
                runProperties38.Append(kern60);
                runProperties38.Append(fontSize39);
                runProperties38.Append(fontSizeComplexScript60);
                Text text38 = new Text();
                text38.Text = "名，";

                run38.Append(runProperties38);
                run38.Append(text38);
                ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run39 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties39 = new RunProperties();
                RunFonts runFonts61 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color40 = new Color() { Val = "000000" };
                Kern kern61 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize40 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "24" };

                runProperties39.Append(runFonts61);
                runProperties39.Append(color40);
                runProperties39.Append(kern61);
                runProperties39.Append(fontSize40);
                runProperties39.Append(fontSizeComplexScript61);
                Text text39 = new Text();
                text39.Text = "均已依";

                run39.Append(runProperties39);
                run39.Append(text39);
                ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run40 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties40 = new RunProperties();
                RunFonts runFonts62 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color41 = new Color() { Val = "000000" };
                Kern kern62 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize41 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "24" };

                runProperties40.Append(runFonts62);
                runProperties40.Append(color41);
                runProperties40.Append(kern62);
                runProperties40.Append(fontSize41);
                runProperties40.Append(fontSizeComplexScript62);
                Text text40 = new Text();
                text40.Text = "計畫實習期滿，並撰寫查核實習心得報告經核";

                run40.Append(runProperties40);
                run40.Append(text40);
                ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

                Run run41 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties41 = new RunProperties();
                RunFonts runFonts63 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color42 = new Color() { Val = "000000" };
                Kern kern63 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize42 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "24" };

                runProperties41.Append(runFonts63);
                runProperties41.Append(color42);
                runProperties41.Append(kern63);
                runProperties41.Append(fontSize42);
                runProperties41.Append(fontSizeComplexScript63);
                Text text41 = new Text();
                text41.Text = "可留卷備查";

                run41.Append(runProperties41);
                run41.Append(text41);
                ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

                Run run42 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties42 = new RunProperties();
                RunFonts runFonts64 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color43 = new Color() { Val = "000000" };
                Kern kern64 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize43 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "24" };

                runProperties42.Append(runFonts64);
                runProperties42.Append(color43);
                runProperties42.Append(kern64);
                runProperties42.Append(fontSize43);
                runProperties42.Append(fontSizeComplexScript64);
                Text text42 = new Text();
                text42.Text = "，依規定核發「查核實習證明書」共";

                run42.Append(runProperties42);
                run42.Append(text42);

                Run run43 = new Run() { RsidRunAddition = "007158A7" };

                RunProperties runProperties43 = new RunProperties();
                RunFonts runFonts65 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color44 = new Color() { Val = "000000" };
                Kern kern65 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize44 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "24" };

                runProperties43.Append(runFonts65);
                runProperties43.Append(color44);
                runProperties43.Append(kern65);
                runProperties43.Append(fontSize44);
                runProperties43.Append(fontSizeComplexScript65);
                Text text43 = new Text();
                text43.Text = dt.Rows[0]["p_num"].ToString();

                run43.Append(runProperties43);
                run43.Append(text43);

                Run run44 = new Run() { RsidRunProperties = "002D5CA1", RsidRunAddition = "002D5CA1" };

                RunProperties runProperties44 = new RunProperties();
                RunFonts runFonts66 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", EastAsia = "標楷體" };
                Color color45 = new Color() { Val = "000000" };
                Kern kern66 = new Kern() { Val = (UInt32Value)0U };
                FontSize fontSize45 = new FontSize() { Val = "31" };
                FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "24" };

                runProperties44.Append(runFonts66);
                runProperties44.Append(color45);
                runProperties44.Append(kern66);
                runProperties44.Append(fontSize45);
                runProperties44.Append(fontSizeComplexScript66);
                Text text44 = new Text();
                text44.Text = "份，請　查照。";

                run44.Append(runProperties44);
                run44.Append(text44);

                paragraph22.Append(paragraphProperties22);
                paragraph22.Append(run35);
                paragraph22.Append(run36);
                paragraph22.Append(run37);
                paragraph22.Append(run38);
                paragraph22.Append(proofError7);
                paragraph22.Append(run39);
                paragraph22.Append(proofError8);
                paragraph22.Append(run40);
                paragraph22.Append(proofError9);
                paragraph22.Append(run41);
                paragraph22.Append(proofError10);
                paragraph22.Append(run42);
                paragraph22.Append(run43);
                paragraph22.Append(run44);

                Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "002D5CA1", RsidRunAdditionDefault = "002D5CA1" };

                ParagraphProperties paragraphProperties23 = new ParagraphProperties();
                AutoSpaceDE autoSpaceDE23 = new AutoSpaceDE() { Val = false };
                AutoSpaceDN autoSpaceDN23 = new AutoSpaceDN() { Val = false };
                AdjustRightIndent adjustRightIndent23 = new AdjustRightIndent() { Val = false };

                ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
                RunFonts runFonts67 = new RunFonts() { Ascii = "標楷體", EastAsia = "標楷體" };
                Kern kern67 = new Kern() { Val = (UInt32Value)0U };
                FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "24" };

                paragraphMarkRunProperties23.Append(runFonts67);
                paragraphMarkRunProperties23.Append(kern67);
                paragraphMarkRunProperties23.Append(fontSizeComplexScript67);

                paragraphProperties23.Append(autoSpaceDE23);
                paragraphProperties23.Append(autoSpaceDN23);
                paragraphProperties23.Append(adjustRightIndent23);
                paragraphProperties23.Append(paragraphMarkRunProperties23);

                paragraph23.Append(paragraphProperties23);

                SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "002D5CA1" };
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
                RsidRoot rsidRoot1 = new RsidRoot() { Val = "002D5CA1" };
                Rsid rsid1 = new Rsid() { Val = "001B650F" };
                Rsid rsid2 = new Rsid() { Val = "002D5CA1" };
                Rsid rsid3 = new Rsid() { Val = "005416E1" };
                Rsid rsid4 = new Rsid() { Val = "007158A7" };

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
                RunFonts runFonts68 = new RunFonts() { ComplexScript = "Times New Roman", AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorEastAsia };
                Kern kern68 = new Kern() { Val = (UInt32Value)2U };
                FontSize fontSize46 = new FontSize() { Val = "24" };
                FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "24" };
                Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

                runPropertiesBaseStyle1.Append(runFonts68);
                runPropertiesBaseStyle1.Append(kern68);
                runPropertiesBaseStyle1.Append(fontSize46);
                runPropertiesBaseStyle1.Append(fontSizeComplexScript68);
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
                RunFonts runFonts69 = new RunFonts() { ComplexScriptTheme = ThemeFontValues.MinorBidi };
                FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "22" };

                styleRunProperties1.Append(runFonts69);
                styleRunProperties1.Append(fontSizeComplexScript69);

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
                Rsid rsid5 = new Rsid() { Val = "002D5CA1" };

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
                FontSize fontSize47 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties2.Append(fontSize47);
                styleRunProperties2.Append(fontSizeComplexScript70);

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
                Rsid rsid6 = new Rsid() { Val = "002D5CA1" };

                StyleRunProperties styleRunProperties3 = new StyleRunProperties();
                RunFonts runFonts70 = new RunFonts() { ComplexScript = "Times New Roman" };
                FontSize fontSize48 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties3.Append(runFonts70);
                styleRunProperties3.Append(fontSize48);
                styleRunProperties3.Append(fontSizeComplexScript71);

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
                Rsid rsid7 = new Rsid() { Val = "002D5CA1" };

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
                FontSize fontSize49 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties4.Append(fontSize49);
                styleRunProperties4.Append(fontSizeComplexScript72);

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
                Rsid rsid8 = new Rsid() { Val = "002D5CA1" };

                StyleRunProperties styleRunProperties5 = new StyleRunProperties();
                RunFonts runFonts71 = new RunFonts() { ComplexScript = "Times New Roman" };
                FontSize fontSize50 = new FontSize() { Val = "20" };
                FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "20" };

                styleRunProperties5.Append(runFonts71);
                styleRunProperties5.Append(fontSize50);
                styleRunProperties5.Append(fontSizeComplexScript73);

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

                Paragraph paragraph24 = new Paragraph() { RsidParagraphAddition = "001B650F", RsidParagraphProperties = "002D5CA1", RsidRunAdditionDefault = "001B650F" };

                Run run45 = new Run();
                SeparatorMark separatorMark1 = new SeparatorMark();

                run45.Append(separatorMark1);

                paragraph24.Append(run45);

                endnote1.Append(paragraph24);

                Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

                Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "001B650F", RsidParagraphProperties = "002D5CA1", RsidRunAdditionDefault = "001B650F" };

                Run run46 = new Run();
                ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

                run46.Append(continuationSeparatorMark1);

                paragraph25.Append(run46);

                endnote2.Append(paragraph25);

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

                Paragraph paragraph26 = new Paragraph() { RsidParagraphAddition = "001B650F", RsidParagraphProperties = "002D5CA1", RsidRunAdditionDefault = "001B650F" };

                Run run47 = new Run();
                SeparatorMark separatorMark2 = new SeparatorMark();

                run47.Append(separatorMark2);

                paragraph26.Append(run47);

                footnote1.Append(paragraph26);

                Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

                Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "001B650F", RsidParagraphProperties = "002D5CA1", RsidRunAdditionDefault = "001B650F" };

                Run run48 = new Run();
                ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

                run48.Append(continuationSeparatorMark2);

                paragraph27.Append(run48);

                footnote2.Append(paragraph27);

                footnotes1.Append(footnote1);
                footnotes1.Append(footnote2);

                footnotesPart1.Footnotes = footnotes1;
            }

            private void SetPackageProperties(OpenXmlPackage document)
            {
                document.PackageProperties.Creator = "Crystal Reports";
                document.PackageProperties.Description = "Powered By Crystal";
                document.PackageProperties.Revision = "2";
                document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2014-02-12T05:31:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
                document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2014-02-12T05:31:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
                document.PackageProperties.LastModifiedBy = "Sony";
            }
        }
}