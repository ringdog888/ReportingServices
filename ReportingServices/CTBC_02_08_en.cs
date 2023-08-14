using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using System;
using System.Data;

namespace CTBC_02_08_OPENXML
{
    public class GeneratedClass_en
    {
         //Data Source
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

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId13");
            GenerateFontTablePart1Content(fontTablePart1);

            CustomXmlPart customXmlPart1 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId3");
            GenerateCustomXmlPart1Content(customXmlPart1);

            CustomXmlPropertiesPart customXmlPropertiesPart1 = customXmlPart1.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart1Content(customXmlPropertiesPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId7");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            FooterPart footerPart1 = mainDocumentPart1.AddNewPart<FooterPart>("rId12");
            GenerateFooterPart1Content(footerPart1);

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

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId15");
            GenerateThemePart1Content(themePart1);

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId10");
            GenerateEndnotesPart1Content(endnotesPart1);

            CustomXmlPart customXmlPart4 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId4");
            GenerateCustomXmlPart4Content(customXmlPart4);

            CustomXmlPropertiesPart customXmlPropertiesPart4 = customXmlPart4.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart4Content(customXmlPropertiesPart4);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId9");
            GenerateFootnotesPart1Content(footnotesPart1);

            WordprocessingPeoplePart wordprocessingPeoplePart1 = mainDocumentPart1.AddNewPart<WordprocessingPeoplePart>("rId14");
            GenerateWordprocessingPeoplePart1Content(wordprocessingPeoplePart1);

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
            totalTime1.Text = "8";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "3";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "259";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "1478";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "12";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "3";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "1734";
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

            Body body1 = new Body();

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "00E659D9", ParagraphId = "102367B6", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a3" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold1 = new Bold();
            BoldComplexScript boldComplexScript1 = new BoldComplexScript();
            FontSize fontSize1 = new FontSize() { Val = "40" };
            Underline underline1 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(bold1);
            paragraphMarkRunProperties1.Append(boldComplexScript1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(underline1);

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold2 = new Bold();
            BoldComplexScript boldComplexScript2 = new BoldComplexScript();
            FontSize fontSize2 = new FontSize() { Val = "40" };
            Underline underline2 = new Underline() { Val = UnderlineValues.Single };

            runProperties1.Append(runFonts2);
            runProperties1.Append(bold2);
            runProperties1.Append(boldComplexScript2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(underline2);
            Text text1 = new Text();
            text1.Text = "Engagement Plan";

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00C90DDE", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "394E5C5A", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "a3" };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize3 = new FontSize() { Val = "40" };
            Underline underline3 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(bold3);
            paragraphMarkRunProperties2.Append(boldComplexScript3);
            paragraphMarkRunProperties2.Append(fontSize3);
            paragraphMarkRunProperties2.Append(underline3);

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            paragraph2.Append(paragraphProperties2);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "10400", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 28, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 28, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 28, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);
            TableLook tableLook1 = new TableLook() { Val = "0000" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableCellMarginDefault1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2410" };
            GridColumn gridColumn2 = new GridColumn() { Width = "7990" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "00E659D9", RsidTableRowProperties = "00E659D9", ParagraphId = "51C97C44", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "0E39BB1F", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };
            Indentation indentation1 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification3 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold4 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties3.Append(runFonts4);
            paragraphMarkRunProperties3.Append(bold4);
            paragraphMarkRunProperties3.Append(fontSize4);

            paragraphProperties3.Append(snapToGrid1);
            paragraphProperties3.Append(indentation1);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run2 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold5 = new Bold();
            FontSize fontSize5 = new FontSize() { Val = "28" };
            Languages languages1 = new Languages() { EastAsia = "zh-HK" };

            runProperties2.Append(runFonts5);
            runProperties2.Append(bold5);
            runProperties2.Append(fontSize5);
            runProperties2.Append(languages1);
            Text text2 = new Text();
            text2.Text = "Audit Project";

            run2.Append(runProperties2);
            run2.Append(text2);

            Run run3 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold6 = new Bold();
            FontSize fontSize6 = new FontSize() { Val = "28" };

            runProperties3.Append(runFonts6);
            runProperties3.Append(bold6);
            runProperties3.Append(fontSize6);
            Text text3 = new Text();
            text3.Text = ":";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run2);
            paragraph3.Append(run3);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph3);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "57CC403E", TextId = "2B27F182" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };
            Indentation indentation2 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification4 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold7 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            Color color1 = new Color() { Val = "0000FF" };
            FontSize fontSize7 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(bold7);
            paragraphMarkRunProperties4.Append(boldComplexScript4);
            paragraphMarkRunProperties4.Append(color1);
            paragraphMarkRunProperties4.Append(fontSize7);
            paragraphMarkRunProperties4.Append(fontSizeComplexScript1);

            paragraphProperties4.Append(snapToGrid2);
            paragraphProperties4.Append(indentation2);
            paragraphProperties4.Append(justification4);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            DeletedRun deletedRun1 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:02:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "0" };

            Run run4 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold8 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            Color color2 = new Color() { Val = "0000FF" };
            FontSize fontSize8 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };

            runProperties4.Append(runFonts8);
            runProperties4.Append(bold8);
            runProperties4.Append(boldComplexScript5);
            runProperties4.Append(color2);
            runProperties4.Append(fontSize8);
            runProperties4.Append(fontSizeComplexScript2);
            DeletedText deletedText1 = new DeletedText();
            deletedText1.Text = "[";

            run4.Append(runProperties4);
            run4.Append(deletedText1);

            Run run5 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold9 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            Color color3 = new Color() { Val = "0000FF" };
            FontSize fontSize9 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };

            runProperties5.Append(runFonts9);
            runProperties5.Append(bold9);
            runProperties5.Append(boldComplexScript6);
            runProperties5.Append(color3);
            runProperties5.Append(fontSize9);
            runProperties5.Append(fontSizeComplexScript3);
            DeletedText deletedText2 = new DeletedText();
            deletedText2.Text = "查程";

            run5.Append(runProperties5);
            run5.Append(deletedText2);

            Run run6 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold10 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            Color color4 = new Color() { Val = "0000FF" };
            FontSize fontSize10 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };

            runProperties6.Append(runFonts10);
            runProperties6.Append(bold10);
            runProperties6.Append(boldComplexScript7);
            runProperties6.Append(color4);
            runProperties6.Append(fontSize10);
            runProperties6.Append(fontSizeComplexScript4);
            DeletedText deletedText3 = new DeletedText();
            deletedText3.Text = "].[";

            run6.Append(runProperties6);
            run6.Append(deletedText3);

            deletedRun1.Append(run4);
            deletedRun1.Append(run5);
            deletedRun1.Append(run6);

            Run run7 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold11 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            Color color5 = new Color() { Val = "0000FF" };
            FontSize fontSize11 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };

            runProperties7.Append(runFonts11);
            runProperties7.Append(bold11);
            runProperties7.Append(boldComplexScript8);
            runProperties7.Append(color5);
            runProperties7.Append(fontSize11);
            runProperties7.Append(fontSizeComplexScript5);
            Text text4 = new Text();
            text4.Text = dt.Rows[0]["planname"].ToString();

            run7.Append(runProperties7);
            run7.Append(text4);

            DeletedRun deletedRun2 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "1" };

            Run run8 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "0096615A" };

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold12 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            Color color6 = new Color() { Val = "0000FF" };
            FontSize fontSize12 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };

            runProperties8.Append(runFonts12);
            runProperties8.Append(bold12);
            runProperties8.Append(boldComplexScript9);
            runProperties8.Append(color6);
            runProperties8.Append(fontSize12);
            runProperties8.Append(fontSizeComplexScript6);
            DeletedText deletedText4 = new DeletedText();
            deletedText4.Text = "_ENG";

            run8.Append(runProperties8);
            run8.Append(deletedText4);

            deletedRun2.Append(run8);

            DeletedRun deletedRun3 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:02:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "2" };

            Run run9 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold13 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            Color color7 = new Color() { Val = "0000FF" };
            FontSize fontSize13 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "28" };

            runProperties9.Append(runFonts13);
            runProperties9.Append(bold13);
            runProperties9.Append(boldComplexScript10);
            runProperties9.Append(color7);
            runProperties9.Append(fontSize13);
            runProperties9.Append(fontSizeComplexScript7);
            DeletedText deletedText5 = new DeletedText();
            deletedText5.Text = "]";

            run9.Append(runProperties9);
            run9.Append(deletedText5);

            deletedRun3.Append(run9);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(deletedRun1);
            paragraph4.Append(run7);
            paragraph4.Append(deletedRun2);
            paragraph4.Append(deletedRun3);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph4);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00E659D9", ParagraphId = "7643CA19", TextId = "77777777" };

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "39B10C34", TextId = "77777777" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            SnapToGrid snapToGrid3 = new SnapToGrid() { Val = false };
            Indentation indentation3 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification5 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold14 = new Bold();
            FontSize fontSize14 = new FontSize() { Val = "28" };
            Languages languages2 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties5.Append(runFonts14);
            paragraphMarkRunProperties5.Append(bold14);
            paragraphMarkRunProperties5.Append(fontSize14);
            paragraphMarkRunProperties5.Append(languages2);

            paragraphProperties5.Append(snapToGrid3);
            paragraphProperties5.Append(indentation3);
            paragraphProperties5.Append(justification5);
            paragraphProperties5.Append(paragraphMarkRunProperties5);

            Run run10 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold15 = new Bold();
            FontSize fontSize15 = new FontSize() { Val = "28" };

            runProperties10.Append(runFonts15);
            runProperties10.Append(bold15);
            runProperties10.Append(fontSize15);
            Text text5 = new Text();
            text5.Text = "Auditee:";

            run10.Append(runProperties10);
            run10.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run10);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph5);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties4.Append(tableCellWidth4);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "582DF604", TextId = "77777777" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SnapToGrid snapToGrid4 = new SnapToGrid() { Val = false };
            Indentation indentation4 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification6 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold16 = new Bold();
            Color color8 = new Color() { Val = "0000FF" };
            FontSize fontSize16 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties6.Append(runFonts16);
            paragraphMarkRunProperties6.Append(bold16);
            paragraphMarkRunProperties6.Append(color8);
            paragraphMarkRunProperties6.Append(fontSize16);

            paragraphProperties6.Append(snapToGrid4);
            paragraphProperties6.Append(indentation4);
            paragraphProperties6.Append(justification6);
            paragraphProperties6.Append(paragraphMarkRunProperties6);

            DeletedRun deletedRun4 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "3" };

            Run run11 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold17 = new Bold();
            Color color9 = new Color() { Val = "0000FF" };
            FontSize fontSize17 = new FontSize() { Val = "28" };

            runProperties11.Append(runFonts17);
            runProperties11.Append(bold17);
            runProperties11.Append(color9);
            runProperties11.Append(fontSize17);
            DeletedText deletedText6 = new DeletedText();
            deletedText6.Text = "UNION([";

            run11.Append(runProperties11);
            run11.Append(deletedText6);

            deletedRun4.Append(run11);

            DeletedRun deletedRun5 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:02:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "4" };

            Run run12 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold18 = new Bold();
            Color color10 = new Color() { Val = "0000FF" };
            FontSize fontSize18 = new FontSize() { Val = "28" };

            runProperties12.Append(runFonts18);
            runProperties12.Append(bold18);
            runProperties12.Append(color10);
            runProperties12.Append(fontSize18);
            DeletedText deletedText7 = new DeletedText();
            deletedText7.Text = "查程";

            run12.Append(runProperties12);
            run12.Append(deletedText7);

            Run run13 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold19 = new Bold();
            Color color11 = new Color() { Val = "0000FF" };
            FontSize fontSize19 = new FontSize() { Val = "28" };

            runProperties13.Append(runFonts19);
            runProperties13.Append(bold19);
            runProperties13.Append(color11);
            runProperties13.Append(fontSize19);
            DeletedText deletedText8 = new DeletedText();
            deletedText8.Text = "].[";

            run13.Append(runProperties13);
            run13.Append(deletedText8);

            deletedRun5.Append(run12);
            deletedRun5.Append(run13);

            Run run14 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold20 = new Bold();
            Color color12 = new Color() { Val = "0000FF" };
            FontSize fontSize20 = new FontSize() { Val = "28" };

            runProperties14.Append(runFonts20);
            runProperties14.Append(bold20);
            runProperties14.Append(color12);
            runProperties14.Append(fontSize20);
            Text text6 = new Text();
            text6.Text = dt.Rows[0]["auditplandept"].ToString();

            run14.Append(runProperties14);
            run14.Append(text6);

            DeletedRun deletedRun6 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "5" };

            Run run15 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "0096615A", RsidRunAddition = "00E659D9" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold21 = new Bold();
            Color color13 = new Color() { Val = "0000FF" };
            FontSize fontSize21 = new FontSize() { Val = "28" };

            runProperties15.Append(runFonts21);
            runProperties15.Append(bold21);
            runProperties15.Append(color13);
            runProperties15.Append(fontSize21);
            DeletedText deletedText9 = new DeletedText();
            deletedText9.Text = "_Eng";

            run15.Append(runProperties15);
            run15.Append(deletedText9);

            Run run16 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "0096615A" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold22 = new Bold();
            Color color14 = new Color() { Val = "0000FF" };
            FontSize fontSize22 = new FontSize() { Val = "28" };

            runProperties16.Append(runFonts22);
            runProperties16.Append(bold22);
            runProperties16.Append(color14);
            runProperties16.Append(fontSize22);
            DeletedText deletedText10 = new DeletedText();
            deletedText10.Text = "]";

            run16.Append(runProperties16);
            run16.Append(deletedText10);

            deletedRun6.Append(run15);
            deletedRun6.Append(run16);

            DeletedRun deletedRun7 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "6" };

            Run run17 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold23 = new Bold();
            Color color15 = new Color() { Val = "0000FF" };
            FontSize fontSize23 = new FontSize() { Val = "28" };

            runProperties17.Append(runFonts23);
            runProperties17.Append(bold23);
            runProperties17.Append(color15);
            runProperties17.Append(fontSize23);
            DeletedText deletedText11 = new DeletedText();
            deletedText11.Text = ")";

            run17.Append(runProperties17);
            run17.Append(deletedText11);

            deletedRun7.Append(run17);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(deletedRun4);
            paragraph6.Append(deletedRun5);
            paragraph6.Append(run14);
            paragraph6.Append(deletedRun6);
            paragraph6.Append(deletedRun7);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph6);

            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00E659D9", ParagraphId = "693B9657", TextId = "77777777" };

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties5.Append(tableCellWidth5);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "21AD0856", TextId = "77777777" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            SnapToGrid snapToGrid5 = new SnapToGrid() { Val = false };
            Indentation indentation5 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification7 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold24 = new Bold();
            FontSize fontSize24 = new FontSize() { Val = "28" };
            Languages languages3 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties7.Append(runFonts24);
            paragraphMarkRunProperties7.Append(bold24);
            paragraphMarkRunProperties7.Append(fontSize24);
            paragraphMarkRunProperties7.Append(languages3);

            paragraphProperties7.Append(snapToGrid5);
            paragraphProperties7.Append(indentation5);
            paragraphProperties7.Append(justification7);
            paragraphProperties7.Append(paragraphMarkRunProperties7);

            Run run18 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold25 = new Bold();
            FontSize fontSize25 = new FontSize() { Val = "28" };
            Languages languages4 = new Languages() { EastAsia = "zh-HK" };

            runProperties18.Append(runFonts25);
            runProperties18.Append(bold25);
            runProperties18.Append(fontSize25);
            runProperties18.Append(languages4);
            Text text7 = new Text();
            text7.Text = "Audit Type";

            run18.Append(runProperties18);
            run18.Append(text7);

            Run run19 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold26 = new Bold();
            FontSize fontSize26 = new FontSize() { Val = "28" };

            runProperties19.Append(runFonts26);
            runProperties19.Append(bold26);
            runProperties19.Append(fontSize26);
            Text text8 = new Text();
            text8.Text = ":";

            run19.Append(runProperties19);
            run19.Append(text8);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run18);
            paragraph7.Append(run19);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph7);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties6.Append(tableCellWidth6);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00E659D9", ParagraphId = "06258A69", TextId = "77777777" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SnapToGrid snapToGrid6 = new SnapToGrid() { Val = false };
            Indentation indentation6 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification8 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts27 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold27 = new Bold();
            Color color16 = new Color() { Val = "0000FF" };
            FontSize fontSize27 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties8.Append(runFonts27);
            paragraphMarkRunProperties8.Append(bold27);
            paragraphMarkRunProperties8.Append(color16);
            paragraphMarkRunProperties8.Append(fontSize27);

            paragraphProperties8.Append(snapToGrid6);
            paragraphProperties8.Append(indentation6);
            paragraphProperties8.Append(justification8);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            DeletedRun deletedRun8 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "7" };

            Run run20 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold28 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            Color color17 = new Color() { Val = "0000FF" };
            FontSize fontSize28 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "28" };

            runProperties20.Append(runFonts28);
            runProperties20.Append(bold28);
            runProperties20.Append(boldComplexScript11);
            runProperties20.Append(color17);
            runProperties20.Append(fontSize28);
            runProperties20.Append(fontSizeComplexScript8);
            DeletedText deletedText12 = new DeletedText();
            deletedText12.Text = "[";

            run20.Append(runProperties20);
            run20.Append(deletedText12);

            Run run21 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold29 = new Bold();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            Color color18 = new Color() { Val = "0000FF" };
            FontSize fontSize29 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "28" };

            runProperties21.Append(runFonts29);
            runProperties21.Append(bold29);
            runProperties21.Append(boldComplexScript12);
            runProperties21.Append(color18);
            runProperties21.Append(fontSize29);
            runProperties21.Append(fontSizeComplexScript9);
            DeletedText deletedText13 = new DeletedText();
            deletedText13.Text = "查程";

            run21.Append(runProperties21);
            run21.Append(deletedText13);

            Run run22 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold30 = new Bold();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            Color color19 = new Color() { Val = "0000FF" };
            FontSize fontSize30 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "28" };

            runProperties22.Append(runFonts30);
            runProperties22.Append(bold30);
            runProperties22.Append(boldComplexScript13);
            runProperties22.Append(color19);
            runProperties22.Append(fontSize30);
            runProperties22.Append(fontSizeComplexScript10);
            DeletedText deletedText14 = new DeletedText();
            deletedText14.Text = "].[";

            run22.Append(runProperties22);
            run22.Append(deletedText14);

            deletedRun8.Append(run20);
            deletedRun8.Append(run21);
            deletedRun8.Append(run22);

            Run run23 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold31 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();
            Color color20 = new Color() { Val = "0000FF" };
            FontSize fontSize31 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "28" };

            runProperties23.Append(runFonts31);
            runProperties23.Append(bold31);
            runProperties23.Append(boldComplexScript14);
            runProperties23.Append(color20);
            runProperties23.Append(fontSize31);
            runProperties23.Append(fontSizeComplexScript11);
            Text text9 = new Text();
            text9.Text = dt.Rows[0]["plantype"].ToString();

            run23.Append(runProperties23);
            run23.Append(text9);

            DeletedRun deletedRun9 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "8" };

            Run run24 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "0096615A" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold32 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            Color color21 = new Color() { Val = "0000FF" };
            FontSize fontSize32 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "28" };

            runProperties24.Append(runFonts32);
            runProperties24.Append(bold32);
            runProperties24.Append(boldComplexScript15);
            runProperties24.Append(color21);
            runProperties24.Append(fontSize32);
            runProperties24.Append(fontSizeComplexScript12);
            DeletedText deletedText15 = new DeletedText();
            deletedText15.Text = "_ENG";

            run24.Append(runProperties24);
            run24.Append(deletedText15);

            deletedRun9.Append(run24);

            DeletedRun deletedRun10 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "9" };

            Run run25 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold33 = new Bold();
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();
            Color color22 = new Color() { Val = "0000FF" };
            FontSize fontSize33 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "28" };

            runProperties25.Append(runFonts33);
            runProperties25.Append(bold33);
            runProperties25.Append(boldComplexScript16);
            runProperties25.Append(color22);
            runProperties25.Append(fontSize33);
            runProperties25.Append(fontSizeComplexScript13);
            DeletedText deletedText16 = new DeletedText();
            deletedText16.Text = "]";

            run25.Append(runProperties25);
            run25.Append(deletedText16);

            deletedRun10.Append(run25);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(deletedRun8);
            paragraph8.Append(run23);
            paragraph8.Append(deletedRun9);
            paragraph8.Append(deletedRun10);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph8);

            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "00E659D9", RsidTableRowProperties = "00E659D9", ParagraphId = "37F0087B", TextId = "77777777" };

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties7.Append(tableCellWidth7);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "2E036A4A", TextId = "77777777" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            SnapToGrid snapToGrid7 = new SnapToGrid() { Val = false };
            Indentation indentation7 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification9 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts34 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold34 = new Bold();
            FontSize fontSize34 = new FontSize() { Val = "28" };
            Languages languages5 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties9.Append(runFonts34);
            paragraphMarkRunProperties9.Append(bold34);
            paragraphMarkRunProperties9.Append(fontSize34);
            paragraphMarkRunProperties9.Append(languages5);

            paragraphProperties9.Append(snapToGrid7);
            paragraphProperties9.Append(indentation7);
            paragraphProperties9.Append(justification9);
            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run26 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold35 = new Bold();
            FontSize fontSize35 = new FontSize() { Val = "28" };
            Languages languages6 = new Languages() { EastAsia = "zh-HK" };

            runProperties26.Append(runFonts35);
            runProperties26.Append(bold35);
            runProperties26.Append(fontSize35);
            runProperties26.Append(languages6);
            Text text10 = new Text();
            text10.Text = "Audit Period:";

            run26.Append(runProperties26);
            run26.Append(text10);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run26);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph9);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties8.Append(tableCellWidth8);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "541154D3", TextId = "4471EC56" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            SnapToGrid snapToGrid8 = new SnapToGrid() { Val = false };
            Indentation indentation8 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification10 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold36 = new Bold();
            Color color23 = new Color() { Val = "0000FF" };
            FontSize fontSize36 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties10.Append(runFonts36);
            paragraphMarkRunProperties10.Append(bold36);
            paragraphMarkRunProperties10.Append(color23);
            paragraphMarkRunProperties10.Append(fontSize36);

            paragraphProperties10.Append(snapToGrid8);
            paragraphProperties10.Append(indentation8);
            paragraphProperties10.Append(justification10);
            paragraphProperties10.Append(paragraphMarkRunProperties10);

            DeletedRun deletedRun11 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "10" };

            Run run27 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold37 = new Bold();
            Color color24 = new Color() { Val = "0000FF" };
            FontSize fontSize37 = new FontSize() { Val = "28" };

            runProperties27.Append(runFonts37);
            runProperties27.Append(bold37);
            runProperties27.Append(color24);
            runProperties27.Append(fontSize37);
            DeletedText deletedText17 = new DeletedText();
            deletedText17.Text = "[";

            run27.Append(runProperties27);
            run27.Append(deletedText17);

            Run run28 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold38 = new Bold();
            Color color25 = new Color() { Val = "0000FF" };
            FontSize fontSize38 = new FontSize() { Val = "28" };

            runProperties28.Append(runFonts38);
            runProperties28.Append(bold38);
            runProperties28.Append(color25);
            runProperties28.Append(fontSize38);
            DeletedText deletedText18 = new DeletedText();
            deletedText18.Text = "查程";

            run28.Append(runProperties28);
            run28.Append(deletedText18);

            Run run29 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold39 = new Bold();
            Color color26 = new Color() { Val = "0000FF" };
            FontSize fontSize39 = new FontSize() { Val = "28" };

            runProperties29.Append(runFonts39);
            runProperties29.Append(bold39);
            runProperties29.Append(color26);
            runProperties29.Append(fontSize39);
            DeletedText deletedText19 = new DeletedText();
            deletedText19.Text = "].[";

            run29.Append(runProperties29);
            run29.Append(deletedText19);

            deletedRun11.Append(run27);
            deletedRun11.Append(run28);
            deletedRun11.Append(run29);

            Run run30 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold40 = new Bold();
            Color color27 = new Color() { Val = "0000FF" };
            FontSize fontSize40 = new FontSize() { Val = "28" };

            runProperties30.Append(runFonts40);
            runProperties30.Append(bold40);
            runProperties30.Append(color27);
            runProperties30.Append(fontSize40);
            Text text11 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["startdate"].ToString(), out DateTime date2))
            {
                text11.Text = date2.ToString("yyyy-MM-dd");
            }
            else
            {
                text11.Text = "";
            }

            run30.Append(runProperties30);
            run30.Append(text11);

            DeletedRun deletedRun12 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "11" };

            Run run31 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold41 = new Bold();
            Color color28 = new Color() { Val = "0000FF" };
            FontSize fontSize41 = new FontSize() { Val = "28" };

            runProperties31.Append(runFonts41);
            runProperties31.Append(bold41);
            runProperties31.Append(color28);
            runProperties31.Append(fontSize41);
            DeletedText deletedText20 = new DeletedText() { Space = SpaceProcessingModeValues.Preserve };
            deletedText20.Text = "] ";

            run31.Append(runProperties31);
            run31.Append(deletedText20);

            deletedRun12.Append(run31);

            Run run32 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold42 = new Bold();
            Color color29 = new Color() { Val = "0000FF" };
            FontSize fontSize42 = new FontSize() { Val = "28" };

            runProperties32.Append(runFonts42);
            runProperties32.Append(bold42);
            runProperties32.Append(color29);
            runProperties32.Append(fontSize42);
            Text text12 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text12.Text = "~ ";

            run32.Append(runProperties32);
            run32.Append(text12);

            DeletedRun deletedRun13 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "12" };

            Run run33 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold43 = new Bold();
            Color color30 = new Color() { Val = "0000FF" };
            FontSize fontSize43 = new FontSize() { Val = "28" };

            runProperties33.Append(runFonts43);
            runProperties33.Append(bold43);
            runProperties33.Append(color30);
            runProperties33.Append(fontSize43);
            DeletedText deletedText21 = new DeletedText();
            deletedText21.Text = "[";

            run33.Append(runProperties33);
            run33.Append(deletedText21);

            Run run34 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold44 = new Bold();
            Color color31 = new Color() { Val = "0000FF" };
            FontSize fontSize44 = new FontSize() { Val = "28" };

            runProperties34.Append(runFonts44);
            runProperties34.Append(bold44);
            runProperties34.Append(color31);
            runProperties34.Append(fontSize44);
            DeletedText deletedText22 = new DeletedText();
            deletedText22.Text = "查程";

            run34.Append(runProperties34);
            run34.Append(deletedText22);

            Run run35 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold45 = new Bold();
            Color color32 = new Color() { Val = "0000FF" };
            FontSize fontSize45 = new FontSize() { Val = "28" };

            runProperties35.Append(runFonts45);
            runProperties35.Append(bold45);
            runProperties35.Append(color32);
            runProperties35.Append(fontSize45);
            DeletedText deletedText23 = new DeletedText();
            deletedText23.Text = "].[";

            run35.Append(runProperties35);
            run35.Append(deletedText23);

            deletedRun13.Append(run33);
            deletedRun13.Append(run34);
            deletedRun13.Append(run35);

            Run run36 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold46 = new Bold();
            Color color33 = new Color() { Val = "0000FF" };
            FontSize fontSize46 = new FontSize() { Val = "28" };

            runProperties36.Append(runFonts46);
            runProperties36.Append(bold46);
            runProperties36.Append(color33);
            runProperties36.Append(fontSize46);
            Text text13 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["enddate"].ToString(), out DateTime date3))
            {
                text13.Text = date3.ToString("yyyy-MM-dd");
            }
            else
            {
                text13.Text = "";
            }
            run36.Append(runProperties36);
            run36.Append(text13);

            DeletedRun deletedRun14 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "13" };

            Run run37 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold47 = new Bold();
            Color color34 = new Color() { Val = "0000FF" };
            FontSize fontSize47 = new FontSize() { Val = "28" };

            runProperties37.Append(runFonts47);
            runProperties37.Append(bold47);
            runProperties37.Append(color34);
            runProperties37.Append(fontSize47);
            DeletedText deletedText24 = new DeletedText();
            deletedText24.Text = "]";

            run37.Append(runProperties37);
            run37.Append(deletedText24);

            deletedRun14.Append(run37);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(deletedRun11);
            paragraph10.Append(run30);
            paragraph10.Append(deletedRun12);
            paragraph10.Append(run32);
            paragraph10.Append(deletedRun13);
            paragraph10.Append(run36);
            paragraph10.Append(deletedRun14);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph10);

            tableRow4.Append(tableCell7);
            tableRow4.Append(tableCell8);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "00E659D9", RsidTableRowProperties = "00E659D9", ParagraphId = "7EABBA80", TextId = "77777777" };

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties9.Append(tableCellWidth9);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "0E822AC2", TextId = "77777777" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            SnapToGrid snapToGrid9 = new SnapToGrid() { Val = false };
            Indentation indentation9 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification11 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts48 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold48 = new Bold();
            FontSize fontSize48 = new FontSize() { Val = "28" };
            Languages languages7 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties11.Append(runFonts48);
            paragraphMarkRunProperties11.Append(bold48);
            paragraphMarkRunProperties11.Append(fontSize48);
            paragraphMarkRunProperties11.Append(languages7);

            paragraphProperties11.Append(snapToGrid9);
            paragraphProperties11.Append(indentation9);
            paragraphProperties11.Append(justification11);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run38 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties38 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold49 = new Bold();
            FontSize fontSize49 = new FontSize() { Val = "28" };
            Languages languages8 = new Languages() { EastAsia = "zh-HK" };

            runProperties38.Append(runFonts49);
            runProperties38.Append(bold49);
            runProperties38.Append(fontSize49);
            runProperties38.Append(languages8);
            Text text14 = new Text();
            text14.Text = "Scope Period:";

            run38.Append(runProperties38);
            run38.Append(text14);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run38);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph11);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties10.Append(tableCellWidth10);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "32AF0050", TextId = "77777777" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            SnapToGrid snapToGrid10 = new SnapToGrid() { Val = false };
            Indentation indentation10 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification12 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts50 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold50 = new Bold();
            Color color35 = new Color() { Val = "0000FF" };
            FontSize fontSize50 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties12.Append(runFonts50);
            paragraphMarkRunProperties12.Append(bold50);
            paragraphMarkRunProperties12.Append(color35);
            paragraphMarkRunProperties12.Append(fontSize50);

            paragraphProperties12.Append(snapToGrid10);
            paragraphProperties12.Append(indentation10);
            paragraphProperties12.Append(justification12);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            DeletedRun deletedRun15 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "14" };

            Run run39 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold51 = new Bold();
            Color color36 = new Color() { Val = "0000FF" };
            FontSize fontSize51 = new FontSize() { Val = "28" };

            runProperties39.Append(runFonts51);
            runProperties39.Append(bold51);
            runProperties39.Append(color36);
            runProperties39.Append(fontSize51);
            DeletedText deletedText25 = new DeletedText();
            deletedText25.Text = "[";

            run39.Append(runProperties39);
            run39.Append(deletedText25);

            Run run40 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold52 = new Bold();
            Color color37 = new Color() { Val = "0000FF" };
            FontSize fontSize52 = new FontSize() { Val = "28" };

            runProperties40.Append(runFonts52);
            runProperties40.Append(bold52);
            runProperties40.Append(color37);
            runProperties40.Append(fontSize52);
            DeletedText deletedText26 = new DeletedText();
            deletedText26.Text = "查程";

            run40.Append(runProperties40);
            run40.Append(deletedText26);

            Run run41 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts53 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold53 = new Bold();
            Color color38 = new Color() { Val = "0000FF" };
            FontSize fontSize53 = new FontSize() { Val = "28" };

            runProperties41.Append(runFonts53);
            runProperties41.Append(bold53);
            runProperties41.Append(color38);
            runProperties41.Append(fontSize53);
            DeletedText deletedText27 = new DeletedText();
            deletedText27.Text = "].[";

            run41.Append(runProperties41);
            run41.Append(deletedText27);

            deletedRun15.Append(run39);
            deletedRun15.Append(run40);
            deletedRun15.Append(run41);

            Run run42 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts54 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold54 = new Bold();
            Color color39 = new Color() { Val = "0000FF" };
            FontSize fontSize54 = new FontSize() { Val = "28" };

            runProperties42.Append(runFonts54);
            runProperties42.Append(bold54);
            runProperties42.Append(color39);
            runProperties42.Append(fontSize54);
            Text text15 = new Text();
            text15.Text = "查核範圍起日";

            run42.Append(runProperties42);
            run42.Append(text15);

            DeletedRun deletedRun16 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "15" };

            Run run43 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold55 = new Bold();
            Color color40 = new Color() { Val = "0000FF" };
            FontSize fontSize55 = new FontSize() { Val = "28" };

            runProperties43.Append(runFonts55);
            runProperties43.Append(bold55);
            runProperties43.Append(color40);
            runProperties43.Append(fontSize55);
            DeletedText deletedText28 = new DeletedText();
            deletedText28.Text = "]";

            run43.Append(runProperties43);
            run43.Append(deletedText28);

            deletedRun16.Append(run43);

            Run run44 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts56 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold56 = new Bold();
            Color color41 = new Color() { Val = "0000FF" };
            FontSize fontSize56 = new FontSize() { Val = "28" };

            runProperties44.Append(runFonts56);
            runProperties44.Append(bold56);
            runProperties44.Append(color41);
            runProperties44.Append(fontSize56);
            Text text16 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text16.Text = " ~ ";

            run44.Append(runProperties44);
            run44.Append(text16);

            DeletedRun deletedRun17 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "16" };

            Run run45 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold57 = new Bold();
            Color color42 = new Color() { Val = "0000FF" };
            FontSize fontSize57 = new FontSize() { Val = "28" };

            runProperties45.Append(runFonts57);
            runProperties45.Append(bold57);
            runProperties45.Append(color42);
            runProperties45.Append(fontSize57);
            DeletedText deletedText29 = new DeletedText();
            deletedText29.Text = "[";

            run45.Append(runProperties45);
            run45.Append(deletedText29);

            Run run46 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold58 = new Bold();
            Color color43 = new Color() { Val = "0000FF" };
            FontSize fontSize58 = new FontSize() { Val = "28" };

            runProperties46.Append(runFonts58);
            runProperties46.Append(bold58);
            runProperties46.Append(color43);
            runProperties46.Append(fontSize58);
            DeletedText deletedText30 = new DeletedText();
            deletedText30.Text = "查程";

            run46.Append(runProperties46);
            run46.Append(deletedText30);

            Run run47 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts59 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold59 = new Bold();
            Color color44 = new Color() { Val = "0000FF" };
            FontSize fontSize59 = new FontSize() { Val = "28" };

            runProperties47.Append(runFonts59);
            runProperties47.Append(bold59);
            runProperties47.Append(color44);
            runProperties47.Append(fontSize59);
            DeletedText deletedText31 = new DeletedText();
            deletedText31.Text = "].[";

            run47.Append(runProperties47);
            run47.Append(deletedText31);

            deletedRun17.Append(run45);
            deletedRun17.Append(run46);
            deletedRun17.Append(run47);

            Run run48 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts60 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold60 = new Bold();
            Color color45 = new Color() { Val = "0000FF" };
            FontSize fontSize60 = new FontSize() { Val = "28" };

            runProperties48.Append(runFonts60);
            runProperties48.Append(bold60);
            runProperties48.Append(color45);
            runProperties48.Append(fontSize60);
            Text text17 = new Text();
            text17.Text = "查核範圍迄日";

            run48.Append(runProperties48);
            run48.Append(text17);

            DeletedRun deletedRun18 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "17" };

            Run run49 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts61 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold61 = new Bold();
            Color color46 = new Color() { Val = "0000FF" };
            FontSize fontSize61 = new FontSize() { Val = "28" };

            runProperties49.Append(runFonts61);
            runProperties49.Append(bold61);
            runProperties49.Append(color46);
            runProperties49.Append(fontSize61);
            DeletedText deletedText32 = new DeletedText();
            deletedText32.Text = "]";

            run49.Append(runProperties49);
            run49.Append(deletedText32);

            deletedRun18.Append(run49);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(deletedRun15);
            paragraph12.Append(run42);
            paragraph12.Append(deletedRun16);
            paragraph12.Append(run44);
            paragraph12.Append(deletedRun17);
            paragraph12.Append(run48);
            paragraph12.Append(deletedRun18);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph12);

            tableRow5.Append(tableCell9);
            tableRow5.Append(tableCell10);

            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "00C7460E", RsidTableRowProperties = "00E659D9", ParagraphId = "21DA0923", TextId = "77777777" };

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties11.Append(tableCellWidth11);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00C7460E", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "1D1EBE9B", TextId = "77777777" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            SnapToGrid snapToGrid11 = new SnapToGrid() { Val = false };
            Indentation indentation11 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification13 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts62 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold62 = new Bold();
            FontSize fontSize62 = new FontSize() { Val = "28" };
            Languages languages9 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties13.Append(runFonts62);
            paragraphMarkRunProperties13.Append(bold62);
            paragraphMarkRunProperties13.Append(fontSize62);
            paragraphMarkRunProperties13.Append(languages9);

            paragraphProperties13.Append(snapToGrid11);
            paragraphProperties13.Append(indentation11);
            paragraphProperties13.Append(justification13);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run50 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties50 = new RunProperties();
            RunFonts runFonts63 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold63 = new Bold();
            FontSize fontSize63 = new FontSize() { Val = "28" };
            Languages languages10 = new Languages() { EastAsia = "zh-HK" };

            runProperties50.Append(runFonts63);
            runProperties50.Append(bold63);
            runProperties50.Append(fontSize63);
            runProperties50.Append(languages10);
            Text text18 = new Text();
            text18.Text = "Auditor in-charge:";

            run50.Append(runProperties50);
            run50.Append(text18);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run50);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph13);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties12.Append(tableCellWidth12);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00C7460E", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C7460E", ParagraphId = "3F830EC2", TextId = "77777777" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            SnapToGrid snapToGrid12 = new SnapToGrid() { Val = false };
            Indentation indentation12 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification14 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts64 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold64 = new Bold();
            Color color47 = new Color() { Val = "0000FF" };
            FontSize fontSize64 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties14.Append(runFonts64);
            paragraphMarkRunProperties14.Append(bold64);
            paragraphMarkRunProperties14.Append(color47);
            paragraphMarkRunProperties14.Append(fontSize64);

            paragraphProperties14.Append(snapToGrid12);
            paragraphProperties14.Append(indentation12);
            paragraphProperties14.Append(justification14);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            DeletedRun deletedRun19 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "18" };

            Run run51 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts65 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold65 = new Bold();
            Color color48 = new Color() { Val = "0000FF" };
            FontSize fontSize65 = new FontSize() { Val = "28" };

            runProperties51.Append(runFonts65);
            runProperties51.Append(bold65);
            runProperties51.Append(color48);
            runProperties51.Append(fontSize65);
            DeletedText deletedText33 = new DeletedText();
            deletedText33.Text = "[";

            run51.Append(runProperties51);
            run51.Append(deletedText33);

            Run run52 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts66 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold66 = new Bold();
            Color color49 = new Color() { Val = "0000FF" };
            FontSize fontSize66 = new FontSize() { Val = "28" };

            runProperties52.Append(runFonts66);
            runProperties52.Append(bold66);
            runProperties52.Append(color49);
            runProperties52.Append(fontSize66);
            DeletedText deletedText34 = new DeletedText();
            deletedText34.Text = "查程";

            run52.Append(runProperties52);
            run52.Append(deletedText34);

            Run run53 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold67 = new Bold();
            Color color50 = new Color() { Val = "0000FF" };
            FontSize fontSize67 = new FontSize() { Val = "28" };

            runProperties53.Append(runFonts67);
            runProperties53.Append(bold67);
            runProperties53.Append(color50);
            runProperties53.Append(fontSize67);
            DeletedText deletedText35 = new DeletedText();
            deletedText35.Text = "].[";

            run53.Append(runProperties53);
            run53.Append(deletedText35);

            deletedRun19.Append(run51);
            deletedRun19.Append(run52);
            deletedRun19.Append(run53);

            Run run54 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts68 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold68 = new Bold();
            Color color51 = new Color() { Val = "0000FF" };
            FontSize fontSize68 = new FontSize() { Val = "28" };

            runProperties54.Append(runFonts68);
            runProperties54.Append(bold68);
            runProperties54.Append(color51);
            runProperties54.Append(fontSize68);
            Text text19 = new Text();
            text19.Text = dt.Rows[0]["leader"].ToString();

            run54.Append(runProperties54);
            run54.Append(text19);

            DeletedRun deletedRun20 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "19" };

            Run run55 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "0096615A", RsidRunAddition = "00F6052A" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold69 = new Bold();
            Color color52 = new Color() { Val = "0000FF" };
            FontSize fontSize69 = new FontSize() { Val = "28" };

            runProperties55.Append(runFonts69);
            runProperties55.Append(bold69);
            runProperties55.Append(color52);
            runProperties55.Append(fontSize69);
            DeletedText deletedText36 = new DeletedText();
            deletedText36.Text = "_Eng";

            run55.Append(runProperties55);
            run55.Append(deletedText36);

            deletedRun20.Append(run55);

            DeletedRun deletedRun21 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "20" };

            Run run56 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts70 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold70 = new Bold();
            Color color53 = new Color() { Val = "0000FF" };
            FontSize fontSize70 = new FontSize() { Val = "28" };

            runProperties56.Append(runFonts70);
            runProperties56.Append(bold70);
            runProperties56.Append(color53);
            runProperties56.Append(fontSize70);
            DeletedText deletedText37 = new DeletedText();
            deletedText37.Text = "]";

            run56.Append(runProperties56);
            run56.Append(deletedText37);

            deletedRun21.Append(run56);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(deletedRun19);
            paragraph14.Append(run54);
            paragraph14.Append(deletedRun20);
            paragraph14.Append(deletedRun21);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph14);

            tableRow6.Append(tableCell11);
            tableRow6.Append(tableCell12);

            TableRow tableRow7 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00E659D9", ParagraphId = "322211A3", TextId = "77777777" };

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties13.Append(tableCellWidth13);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "449D4354", TextId = "77777777" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            SnapToGrid snapToGrid13 = new SnapToGrid() { Val = false };
            Indentation indentation13 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification15 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts71 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold71 = new Bold();
            FontSize fontSize71 = new FontSize() { Val = "28" };
            Languages languages11 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties15.Append(runFonts71);
            paragraphMarkRunProperties15.Append(bold71);
            paragraphMarkRunProperties15.Append(fontSize71);
            paragraphMarkRunProperties15.Append(languages11);

            paragraphProperties15.Append(snapToGrid13);
            paragraphProperties15.Append(indentation13);
            paragraphProperties15.Append(justification15);
            paragraphProperties15.Append(paragraphMarkRunProperties15);

            Run run57 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts72 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold72 = new Bold();
            FontSize fontSize72 = new FontSize() { Val = "28" };
            Languages languages12 = new Languages() { EastAsia = "zh-HK" };

            runProperties57.Append(runFonts72);
            runProperties57.Append(bold72);
            runProperties57.Append(fontSize72);
            runProperties57.Append(languages12);
            Text text20 = new Text();
            text20.Text = "Auditor:";

            run57.Append(runProperties57);
            run57.Append(text20);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run57);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph15);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties14.Append(tableCellWidth14);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "008625F1", ParagraphId = "602CCF85", TextId = "77777777" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            SnapToGrid snapToGrid14 = new SnapToGrid() { Val = false };
            Indentation indentation14 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification16 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            RunFonts runFonts73 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold73 = new Bold();
            Color color54 = new Color() { Val = "0000FF" };
            FontSize fontSize73 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties16.Append(runFonts73);
            paragraphMarkRunProperties16.Append(bold73);
            paragraphMarkRunProperties16.Append(color54);
            paragraphMarkRunProperties16.Append(fontSize73);

            paragraphProperties16.Append(snapToGrid14);
            paragraphProperties16.Append(indentation14);
            paragraphProperties16.Append(justification16);
            paragraphProperties16.Append(paragraphMarkRunProperties16);

            DeletedRun deletedRun22 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "21" };

            Run run58 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts74 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold74 = new Bold();
            Color color55 = new Color() { Val = "0000FF" };
            FontSize fontSize74 = new FontSize() { Val = "28" };

            runProperties58.Append(runFonts74);
            runProperties58.Append(bold74);
            runProperties58.Append(color55);
            runProperties58.Append(fontSize74);
            DeletedText deletedText38 = new DeletedText();
            deletedText38.Text = "UNION([";

            run58.Append(runProperties58);
            run58.Append(deletedText38);

            Run run59 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts75 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold75 = new Bold();
            Color color56 = new Color() { Val = "0000FF" };
            FontSize fontSize75 = new FontSize() { Val = "28" };

            runProperties59.Append(runFonts75);
            runProperties59.Append(bold75);
            runProperties59.Append(color56);
            runProperties59.Append(fontSize75);
            DeletedText deletedText39 = new DeletedText();
            deletedText39.Text = "查程工作分配";

            run59.Append(runProperties59);
            run59.Append(deletedText39);

            Run run60 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties60 = new RunProperties();
            RunFonts runFonts76 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold76 = new Bold();
            Color color57 = new Color() { Val = "0000FF" };
            FontSize fontSize76 = new FontSize() { Val = "28" };

            runProperties60.Append(runFonts76);
            runProperties60.Append(bold76);
            runProperties60.Append(color57);
            runProperties60.Append(fontSize76);
            DeletedText deletedText40 = new DeletedText();
            deletedText40.Text = "].[";

            run60.Append(runProperties60);
            run60.Append(deletedText40);

            deletedRun22.Append(run58);
            deletedRun22.Append(run59);
            deletedRun22.Append(run60);

            Run run61 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts77 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold77 = new Bold();
            Color color58 = new Color() { Val = "0000FF" };
            FontSize fontSize77 = new FontSize() { Val = "28" };

            runProperties61.Append(runFonts77);
            runProperties61.Append(bold77);
            runProperties61.Append(color58);
            runProperties61.Append(fontSize77);
            Text text21 = new Text();
            text21.Text = dt.Rows[0]["Member"].ToString();

            run61.Append(runProperties61);
            run61.Append(text21);

            DeletedRun deletedRun23 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "22" };

            Run run62 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "0096615A", RsidRunAddition = "00F6052A" };

            RunProperties runProperties62 = new RunProperties();
            RunFonts runFonts78 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold78 = new Bold();
            Color color59 = new Color() { Val = "0000FF" };
            FontSize fontSize78 = new FontSize() { Val = "28" };

            runProperties62.Append(runFonts78);
            runProperties62.Append(bold78);
            runProperties62.Append(color59);
            runProperties62.Append(fontSize78);
            DeletedText deletedText41 = new DeletedText();
            deletedText41.Text = "_Eng";

            run62.Append(runProperties62);
            run62.Append(deletedText41);

            deletedRun23.Append(run62);

            DeletedRun deletedRun24 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "23" };

            Run run63 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties63 = new RunProperties();
            RunFonts runFonts79 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold79 = new Bold();
            Color color60 = new Color() { Val = "0000FF" };
            FontSize fontSize79 = new FontSize() { Val = "28" };

            runProperties63.Append(runFonts79);
            runProperties63.Append(bold79);
            runProperties63.Append(color60);
            runProperties63.Append(fontSize79);
            DeletedText deletedText42 = new DeletedText();
            deletedText42.Text = "])";

            run63.Append(runProperties63);
            run63.Append(deletedText42);

            deletedRun24.Append(run63);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(deletedRun22);
            paragraph16.Append(run61);
            paragraph16.Append(deletedRun23);
            paragraph16.Append(deletedRun24);

            tableCell14.Append(tableCellProperties14);
            tableCell14.Append(paragraph16);

            tableRow7.Append(tableCell13);
            tableRow7.Append(tableCell14);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            table1.Append(tableRow4);
            table1.Append(tableRow5);
            table1.Append(tableRow6);
            table1.Append(tableRow7);

            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "600B5867", TextId = "77777777" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            RunFonts runFonts80 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties17.Append(runFonts80);

            paragraphProperties17.Append(paragraphMarkRunProperties17);

            paragraph17.Append(paragraphProperties17);

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00F6052A", RsidRunAdditionDefault = "00F6052A", ParagraphId = "2E4B5E7B", TextId = "77777777" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId1 = new NumberingId() { Val = 15 };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SnapToGrid snapToGrid15 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold80 = new Bold();
            BoldComplexScript boldComplexScript17 = new BoldComplexScript();
            FontSize fontSize80 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties18.Append(runFonts81);
            paragraphMarkRunProperties18.Append(bold80);
            paragraphMarkRunProperties18.Append(boldComplexScript17);
            paragraphMarkRunProperties18.Append(fontSize80);

            paragraphProperties18.Append(numberingProperties1);
            paragraphProperties18.Append(snapToGrid15);
            paragraphProperties18.Append(paragraphMarkRunProperties18);

            Run run64 = new Run();

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts82 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold81 = new Bold();
            BoldComplexScript boldComplexScript18 = new BoldComplexScript();
            FontSize fontSize81 = new FontSize() { Val = "28" };

            runProperties64.Append(runFonts82);
            runProperties64.Append(bold81);
            runProperties64.Append(boldComplexScript18);
            runProperties64.Append(fontSize81);
            Text text22 = new Text();
            text22.Text = "A";

            run64.Append(runProperties64);
            run64.Append(text22);

            Run run65 = new Run() { RsidRunProperties = "00F6052A" };

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts83 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold82 = new Bold();
            BoldComplexScript boldComplexScript19 = new BoldComplexScript();
            FontSize fontSize82 = new FontSize() { Val = "28" };

            runProperties65.Append(runFonts83);
            runProperties65.Append(bold82);
            runProperties65.Append(boldComplexScript19);
            runProperties65.Append(fontSize82);
            Text text23 = new Text();
            text23.Text = "ssignments";

            run65.Append(runProperties65);
            run65.Append(text23);

            Run run66 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts84 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold83 = new Bold();
            BoldComplexScript boldComplexScript20 = new BoldComplexScript();
            FontSize fontSize83 = new FontSize() { Val = "28" };

            runProperties66.Append(runFonts84);
            runProperties66.Append(bold83);
            runProperties66.Append(boldComplexScript20);
            runProperties66.Append(fontSize83);
            Text text24 = new Text();
            text24.Text = ":";

            run66.Append(runProperties66);
            run66.Append(text24);

            paragraph18.Append(paragraphProperties18);
            paragraph18.Append(run64);
            paragraph18.Append(run65);
            paragraph18.Append(run66);

            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "9214", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 595, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)8U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 28, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 28, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);
            TableLook tableLook2 = new TableLook() { Val = "0000" };

            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableIndentation2);
            tableProperties2.Append(tableBorders1);
            tableProperties2.Append(tableCellMarginDefault2);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn3 = new GridColumn() { Width = "1881" };
            GridColumn gridColumn4 = new GridColumn() { Width = "4425" };
            GridColumn gridColumn5 = new GridColumn() { Width = "2908" };

            tableGrid2.Append(gridColumn3);
            tableGrid2.Append(gridColumn4);
            tableGrid2.Append(gridColumn5);

            TableRow tableRow8 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00193FEF", ParagraphId = "454E3D7D", TextId = "77777777" };

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };

            tableCellProperties15.Append(tableCellWidth15);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00F6052A", ParagraphId = "1C746037", TextId = "77777777" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            SnapToGrid snapToGrid16 = new SnapToGrid() { Val = false };
            Indentation indentation15 = new Indentation() { Start = "-24", StartCharacters = -11, Hanging = "2" };
            Justification justification17 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts85 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize84 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties19.Append(runFonts85);
            paragraphMarkRunProperties19.Append(fontSize84);

            paragraphProperties19.Append(snapToGrid16);
            paragraphProperties19.Append(indentation15);
            paragraphProperties19.Append(justification17);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run67 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties67 = new RunProperties();
            RunFonts runFonts86 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize85 = new FontSize() { Val = "28" };

            runProperties67.Append(runFonts86);
            runProperties67.Append(fontSize85);
            Text text25 = new Text();
            text25.Text = "N";

            run67.Append(runProperties67);
            run67.Append(text25);

            Run run68 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts87 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize86 = new FontSize() { Val = "28" };

            runProperties68.Append(runFonts87);
            runProperties68.Append(fontSize86);
            Text text26 = new Text();
            text26.Text = "o.";

            run68.Append(runProperties68);
            run68.Append(text26);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run67);
            paragraph19.Append(run68);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph19);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };

            tableCellProperties16.Append(tableCellWidth16);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0034608C", RsidRunAdditionDefault = "00DE479D", ParagraphId = "379BD50E", TextId = "77777777" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            SnapToGrid snapToGrid17 = new SnapToGrid() { Val = false };
            Justification justification18 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize87 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties20.Append(runFonts88);
            paragraphMarkRunProperties20.Append(fontSize87);

            paragraphProperties20.Append(snapToGrid17);
            paragraphProperties20.Append(justification18);
            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run69 = new Run();

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts89 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize88 = new FontSize() { Val = "28" };

            runProperties69.Append(runFonts89);
            runProperties69.Append(fontSize88);
            Text text27 = new Text();
            text27.Text = "Audit Subject";

            run69.Append(runProperties69);
            run69.Append(text27);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run69);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph20);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };

            tableCellProperties17.Append(tableCellWidth17);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00F6052A", ParagraphId = "704940F7", TextId = "77777777" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            SnapToGrid snapToGrid18 = new SnapToGrid() { Val = false };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts90 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize89 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties21.Append(runFonts90);
            paragraphMarkRunProperties21.Append(fontSize89);

            paragraphProperties21.Append(snapToGrid18);
            paragraphProperties21.Append(justification19);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run70 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties70 = new RunProperties();
            RunFonts runFonts91 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize90 = new FontSize() { Val = "28" };

            runProperties70.Append(runFonts91);
            runProperties70.Append(fontSize90);
            Text text28 = new Text();
            text28.Text = "A";

            run70.Append(runProperties70);
            run70.Append(text28);

            Run run71 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties71 = new RunProperties();
            RunFonts runFonts92 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize91 = new FontSize() { Val = "28" };

            runProperties71.Append(runFonts92);
            runProperties71.Append(fontSize91);
            Text text29 = new Text();
            text29.Text = "uditor";

            run71.Append(runProperties71);
            run71.Append(text29);

            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run70);
            paragraph21.Append(run71);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph21);

            tableRow8.Append(tableCell15);
            tableRow8.Append(tableCell16);
            tableRow8.Append(tableCell17);

            TableRow tableRow9 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00193FEF", ParagraphId = "5932E449", TextId = "77777777" };

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellVerticalAlignment1);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "6599C4B1", TextId = "77777777" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            SnapToGrid snapToGrid19 = new SnapToGrid() { Val = false };
            Justification justification20 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties22.Append(runFonts93);

            paragraphProperties22.Append(snapToGrid19);
            paragraphProperties22.Append(justification20);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            DeletedRun deletedRun25 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "24" };

            Run run72 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties72 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts94 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold84 = new Bold();
            BoldComplexScript boldComplexScript21 = new BoldComplexScript();
            Color color61 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "28" };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties72.Append(runStyle1);
            runProperties72.Append(runFonts94);
            runProperties72.Append(bold84);
            runProperties72.Append(boldComplexScript21);
            runProperties72.Append(color61);
            runProperties72.Append(fontSizeComplexScript14);
            runProperties72.Append(shading1);
            DeletedText deletedText43 = new DeletedText();
            deletedText43.Text = "[";

            run72.Append(runProperties72);
            run72.Append(deletedText43);

            Run run73 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties73 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts95 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold85 = new Bold();
            BoldComplexScript boldComplexScript22 = new BoldComplexScript();
            Color color62 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "28" };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages13 = new Languages() { EastAsia = "zh-HK" };

            runProperties73.Append(runStyle2);
            runProperties73.Append(runFonts95);
            runProperties73.Append(bold85);
            runProperties73.Append(boldComplexScript22);
            runProperties73.Append(color62);
            runProperties73.Append(fontSizeComplexScript15);
            runProperties73.Append(shading2);
            runProperties73.Append(languages13);
            DeletedText deletedText44 = new DeletedText();
            deletedText44.Text = "查程工作分配";

            run73.Append(runProperties73);
            run73.Append(deletedText44);

            Run run74 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties74 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts96 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold86 = new Bold();
            BoldComplexScript boldComplexScript23 = new BoldComplexScript();
            Color color63 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "28" };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties74.Append(runStyle3);
            runProperties74.Append(runFonts96);
            runProperties74.Append(bold86);
            runProperties74.Append(boldComplexScript23);
            runProperties74.Append(color63);
            runProperties74.Append(fontSizeComplexScript16);
            runProperties74.Append(shading3);
            DeletedText deletedText45 = new DeletedText();
            deletedText45.Text = "].[";

            run74.Append(runProperties74);
            run74.Append(deletedText45);

            deletedRun25.Append(run72);
            deletedRun25.Append(run73);
            deletedRun25.Append(run74);

            Run run75 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties75 = new RunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts97 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold87 = new Bold();
            BoldComplexScript boldComplexScript24 = new BoldComplexScript();
            Color color64 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "28" };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages14 = new Languages() { EastAsia = "zh-HK" };

            runProperties75.Append(runStyle4);
            runProperties75.Append(runFonts97);
            runProperties75.Append(bold87);
            runProperties75.Append(boldComplexScript24);
            runProperties75.Append(color64);
            runProperties75.Append(fontSizeComplexScript17);
            runProperties75.Append(shading4);
            runProperties75.Append(languages14);
            Text text30 = new Text();
            //text30.Text = dt.Rows[0]["subcode"].ToString();

            run75.Append(runProperties75);
            run75.Append(text30);

            DeletedRun deletedRun26 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "25" };

            Run run76 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties76 = new RunProperties();
            RunStyle runStyle5 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts98 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold88 = new Bold();
            BoldComplexScript boldComplexScript25 = new BoldComplexScript();
            Color color65 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "28" };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties76.Append(runStyle5);
            runProperties76.Append(runFonts98);
            runProperties76.Append(bold88);
            runProperties76.Append(boldComplexScript25);
            runProperties76.Append(color65);
            runProperties76.Append(fontSizeComplexScript18);
            runProperties76.Append(shading5);
            DeletedText deletedText46 = new DeletedText();
            deletedText46.Text = "]";

            run76.Append(runProperties76);
            run76.Append(deletedText46);

            deletedRun26.Append(run76);

            paragraph22.Append(paragraphProperties22);
            paragraph22.Append(deletedRun25);
            paragraph22.Append(run75);
            paragraph22.Append(deletedRun26);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph22);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellVerticalAlignment2);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "571BA3CC", TextId = "4540A4EF" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            SnapToGrid snapToGrid20 = new SnapToGrid() { Val = false };
            Indentation indentation16 = new Indentation() { FirstLine = "240", FirstLineChars = 100 };
            Justification justification21 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts99 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties23.Append(runFonts99);

            paragraphProperties23.Append(snapToGrid20);
            paragraphProperties23.Append(indentation16);
            paragraphProperties23.Append(justification21);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            DeletedRun deletedRun27 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "26" };

            Run run77 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties77 = new RunProperties();
            RunStyle runStyle6 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts100 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold89 = new Bold();
            BoldComplexScript boldComplexScript26 = new BoldComplexScript();
            Color color66 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "28" };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties77.Append(runStyle6);
            runProperties77.Append(runFonts100);
            runProperties77.Append(bold89);
            runProperties77.Append(boldComplexScript26);
            runProperties77.Append(color66);
            runProperties77.Append(fontSizeComplexScript19);
            runProperties77.Append(shading6);
            DeletedText deletedText47 = new DeletedText();
            deletedText47.Text = "[";

            run77.Append(runProperties77);
            run77.Append(deletedText47);

            Run run78 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties78 = new RunProperties();
            RunStyle runStyle7 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts101 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold90 = new Bold();
            BoldComplexScript boldComplexScript27 = new BoldComplexScript();
            Color color67 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "28" };
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages15 = new Languages() { EastAsia = "zh-HK" };

            runProperties78.Append(runStyle7);
            runProperties78.Append(runFonts101);
            runProperties78.Append(bold90);
            runProperties78.Append(boldComplexScript27);
            runProperties78.Append(color67);
            runProperties78.Append(fontSizeComplexScript20);
            runProperties78.Append(shading7);
            runProperties78.Append(languages15);
            DeletedText deletedText48 = new DeletedText();
            deletedText48.Text = "查程工作分配";

            run78.Append(runProperties78);
            run78.Append(deletedText48);

            Run run79 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties79 = new RunProperties();
            RunStyle runStyle8 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts102 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold91 = new Bold();
            BoldComplexScript boldComplexScript28 = new BoldComplexScript();
            Color color68 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "28" };
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties79.Append(runStyle8);
            runProperties79.Append(runFonts102);
            runProperties79.Append(bold91);
            runProperties79.Append(boldComplexScript28);
            runProperties79.Append(color68);
            runProperties79.Append(fontSizeComplexScript21);
            runProperties79.Append(shading8);
            DeletedText deletedText49 = new DeletedText();
            deletedText49.Text = "].[";

            run79.Append(runProperties79);
            run79.Append(deletedText49);

            deletedRun27.Append(run77);
            deletedRun27.Append(run78);
            deletedRun27.Append(run79);

            Run run80 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties80 = new RunProperties();
            RunStyle runStyle9 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts103 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold92 = new Bold();
            BoldComplexScript boldComplexScript29 = new BoldComplexScript();
            Color color69 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "28" };
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages16 = new Languages() { EastAsia = "zh-HK" };

            runProperties80.Append(runStyle9);
            runProperties80.Append(runFonts103);
            runProperties80.Append(bold92);
            runProperties80.Append(boldComplexScript29);
            runProperties80.Append(color69);
            runProperties80.Append(fontSizeComplexScript22);
            runProperties80.Append(shading9);
            runProperties80.Append(languages16);
            Text text31 = new Text();
            text31.Text = "subname";

            run80.Append(runProperties80);
            run80.Append(text31);

            DeletedRun deletedRun28 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "27" };

            Run run81 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "0096615A", RsidRunAddition = "00F6052A" };

            RunProperties runProperties81 = new RunProperties();
            RunStyle runStyle10 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts104 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold93 = new Bold();
            BoldComplexScript boldComplexScript30 = new BoldComplexScript();
            Color color70 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "28" };
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties81.Append(runStyle10);
            runProperties81.Append(runFonts104);
            runProperties81.Append(bold93);
            runProperties81.Append(boldComplexScript30);
            runProperties81.Append(color70);
            runProperties81.Append(fontSizeComplexScript23);
            runProperties81.Append(shading10);
            DeletedText deletedText50 = new DeletedText();
            deletedText50.Text = "_Eng";

            run81.Append(runProperties81);
            run81.Append(deletedText50);

            deletedRun28.Append(run81);

            DeletedRun deletedRun29 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "28" };

            Run run82 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties82 = new RunProperties();
            RunStyle runStyle11 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts105 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold94 = new Bold();
            BoldComplexScript boldComplexScript31 = new BoldComplexScript();
            Color color71 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "28" };
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties82.Append(runStyle11);
            runProperties82.Append(runFonts105);
            runProperties82.Append(bold94);
            runProperties82.Append(boldComplexScript31);
            runProperties82.Append(color71);
            runProperties82.Append(fontSizeComplexScript24);
            runProperties82.Append(shading11);
            DeletedText deletedText51 = new DeletedText();
            deletedText51.Text = "]";

            run82.Append(runProperties82);
            run82.Append(deletedText51);

            deletedRun29.Append(run82);

            paragraph23.Append(paragraphProperties23);
            paragraph23.Append(deletedRun27);
            paragraph23.Append(run80);
            paragraph23.Append(deletedRun28);
            paragraph23.Append(deletedRun29);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph23);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(tableCellVerticalAlignment3);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "5A67BFB8", TextId = "639840B7" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            SnapToGrid snapToGrid21 = new SnapToGrid() { Val = false };
            Justification justification22 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties24.Append(runFonts106);

            paragraphProperties24.Append(snapToGrid21);
            paragraphProperties24.Append(justification22);
            paragraphProperties24.Append(paragraphMarkRunProperties24);

            DeletedRun deletedRun30 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "29" };

            Run run83 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties83 = new RunProperties();
            RunFonts runFonts107 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold95 = new Bold();
            Color color72 = new Color() { Val = "0000FF" };

            runProperties83.Append(runFonts107);
            runProperties83.Append(bold95);
            runProperties83.Append(color72);
            DeletedText deletedText52 = new DeletedText();
            deletedText52.Text = "[";

            run83.Append(runProperties83);
            run83.Append(deletedText52);

            Run run84 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties84 = new RunProperties();
            RunFonts runFonts108 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold96 = new Bold();
            Color color73 = new Color() { Val = "0000FF" };

            runProperties84.Append(runFonts108);
            runProperties84.Append(bold96);
            runProperties84.Append(color73);
            DeletedText deletedText53 = new DeletedText();
            deletedText53.Text = "查程工作分配";

            run84.Append(runProperties84);
            run84.Append(deletedText53);

            Run run85 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties85 = new RunProperties();
            RunFonts runFonts109 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold97 = new Bold();
            Color color74 = new Color() { Val = "0000FF" };

            runProperties85.Append(runFonts109);
            runProperties85.Append(bold97);
            runProperties85.Append(color74);
            DeletedText deletedText54 = new DeletedText();
            deletedText54.Text = "].[";

            run85.Append(runProperties85);
            run85.Append(deletedText54);

            deletedRun30.Append(run83);
            deletedRun30.Append(run84);
            deletedRun30.Append(run85);

            Run run86 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties86 = new RunProperties();
            RunFonts runFonts110 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold98 = new Bold();
            Color color75 = new Color() { Val = "0000FF" };

            runProperties86.Append(runFonts110);
            runProperties86.Append(bold98);
            runProperties86.Append(color75);
            Text text32 = new Text();
            text32.Text = dt.Rows[0]["Member"].ToString();

            run86.Append(runProperties86);
            run86.Append(text32);

            DeletedRun deletedRun31 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "30" };

            Run run87 = new Run() { RsidRunProperties = "00917606", RsidRunDeletion = "0096615A", RsidRunAddition = "00F6052A" };

            RunProperties runProperties87 = new RunProperties();
            RunStyle runStyle12 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts111 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold99 = new Bold();
            BoldComplexScript boldComplexScript32 = new BoldComplexScript();
            Color color76 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "28" };
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties87.Append(runStyle12);
            runProperties87.Append(runFonts111);
            runProperties87.Append(bold99);
            runProperties87.Append(boldComplexScript32);
            runProperties87.Append(color76);
            runProperties87.Append(fontSizeComplexScript25);
            runProperties87.Append(shading12);
            DeletedText deletedText55 = new DeletedText();
            deletedText55.Text = "_Eng";

            run87.Append(runProperties87);
            run87.Append(deletedText55);

            deletedRun31.Append(run87);

            DeletedRun deletedRun32 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "31" };

            Run run88 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties88 = new RunProperties();
            RunFonts runFonts112 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold100 = new Bold();
            Color color77 = new Color() { Val = "0000FF" };

            runProperties88.Append(runFonts112);
            runProperties88.Append(bold100);
            runProperties88.Append(color77);
            DeletedText deletedText56 = new DeletedText();
            deletedText56.Text = "]";

            run88.Append(runProperties88);
            run88.Append(deletedText56);

            deletedRun32.Append(run88);

            paragraph24.Append(paragraphProperties24);
            paragraph24.Append(deletedRun30);
            paragraph24.Append(run86);
            paragraph24.Append(deletedRun31);
            paragraph24.Append(deletedRun32);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph24);

            tableRow9.Append(tableCell18);
            tableRow9.Append(tableCell19);
            tableRow9.Append(tableCell20);

            TableRow tableRow10 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "00EB5D9F", RsidTableRowProperties = "00193FEF", ParagraphId = "3886A00B", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)80U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(tableCellVerticalAlignment4);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00EB5D9F", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "106CBAAF", TextId = "77777777" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            SnapToGrid snapToGrid22 = new SnapToGrid() { Val = false };
            Justification justification23 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            RunFonts runFonts113 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties25.Append(runFonts113);

            paragraphProperties25.Append(snapToGrid22);
            paragraphProperties25.Append(justification23);
            paragraphProperties25.Append(paragraphMarkRunProperties25);

            DeletedRun deletedRun33 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "32" };

            Run run89 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties89 = new RunProperties();
            RunStyle runStyle13 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts114 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold101 = new Bold();
            BoldComplexScript boldComplexScript33 = new BoldComplexScript();
            Color color78 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "28" };
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties89.Append(runStyle13);
            runProperties89.Append(runFonts114);
            runProperties89.Append(bold101);
            runProperties89.Append(boldComplexScript33);
            runProperties89.Append(color78);
            runProperties89.Append(fontSizeComplexScript26);
            runProperties89.Append(shading13);
            DeletedText deletedText57 = new DeletedText();
            deletedText57.Text = "[";

            run89.Append(runProperties89);
            run89.Append(deletedText57);

            Run run90 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties90 = new RunProperties();
            RunStyle runStyle14 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts115 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold102 = new Bold();
            BoldComplexScript boldComplexScript34 = new BoldComplexScript();
            Color color79 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "28" };
            Shading shading14 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages17 = new Languages() { EastAsia = "zh-HK" };

            runProperties90.Append(runStyle14);
            runProperties90.Append(runFonts115);
            runProperties90.Append(bold102);
            runProperties90.Append(boldComplexScript34);
            runProperties90.Append(color79);
            runProperties90.Append(fontSizeComplexScript27);
            runProperties90.Append(shading14);
            runProperties90.Append(languages17);
            DeletedText deletedText58 = new DeletedText();
            deletedText58.Text = "查程工作分配";

            run90.Append(runProperties90);
            run90.Append(deletedText58);

            Run run91 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties91 = new RunProperties();
            RunStyle runStyle15 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts116 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold103 = new Bold();
            BoldComplexScript boldComplexScript35 = new BoldComplexScript();
            Color color80 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "28" };
            Shading shading15 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties91.Append(runStyle15);
            runProperties91.Append(runFonts116);
            runProperties91.Append(bold103);
            runProperties91.Append(boldComplexScript35);
            runProperties91.Append(color80);
            runProperties91.Append(fontSizeComplexScript28);
            runProperties91.Append(shading15);
            DeletedText deletedText59 = new DeletedText();
            deletedText59.Text = "].[";

            run91.Append(runProperties91);
            run91.Append(deletedText59);

            deletedRun33.Append(run89);
            deletedRun33.Append(run90);
            deletedRun33.Append(run91);

            Run run92 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties92 = new RunProperties();
            RunStyle runStyle16 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts117 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold104 = new Bold();
            BoldComplexScript boldComplexScript36 = new BoldComplexScript();
            Color color81 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "28" };
            Shading shading16 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages18 = new Languages() { EastAsia = "zh-HK" };

            runProperties92.Append(runStyle16);
            runProperties92.Append(runFonts117);
            runProperties92.Append(bold104);
            runProperties92.Append(boldComplexScript36);
            runProperties92.Append(color81);
            runProperties92.Append(fontSizeComplexScript29);
            runProperties92.Append(shading16);
            runProperties92.Append(languages18);
            Text text33 = new Text();
            //text33.Text = dt.Rows[0]["subcode"].ToString();

            run92.Append(runProperties92);
            run92.Append(text33);

            DeletedRun deletedRun34 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:03:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "33" };

            Run run93 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties93 = new RunProperties();
            RunStyle runStyle17 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts118 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold105 = new Bold();
            BoldComplexScript boldComplexScript37 = new BoldComplexScript();
            Color color82 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "28" };
            Shading shading17 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties93.Append(runStyle17);
            runProperties93.Append(runFonts118);
            runProperties93.Append(bold105);
            runProperties93.Append(boldComplexScript37);
            runProperties93.Append(color82);
            runProperties93.Append(fontSizeComplexScript30);
            runProperties93.Append(shading17);
            DeletedText deletedText60 = new DeletedText();
            deletedText60.Text = "]";

            run93.Append(runProperties93);
            run93.Append(deletedText60);

            deletedRun34.Append(run93);

            paragraph25.Append(paragraphProperties25);
            paragraph25.Append(deletedRun33);
            paragraph25.Append(run92);
            paragraph25.Append(deletedRun34);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph25);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(tableCellVerticalAlignment5);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00EB5D9F", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "36429036", TextId = "0C598A08" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            SnapToGrid snapToGrid23 = new SnapToGrid() { Val = false };
            Indentation indentation17 = new Indentation() { FirstLine = "240", FirstLineChars = 100 };
            Justification justification24 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            RunFonts runFonts119 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties26.Append(runFonts119);

            paragraphProperties26.Append(snapToGrid23);
            paragraphProperties26.Append(indentation17);
            paragraphProperties26.Append(justification24);
            paragraphProperties26.Append(paragraphMarkRunProperties26);

            DeletedRun deletedRun35 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "34" };

            Run run94 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties94 = new RunProperties();
            RunStyle runStyle18 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts120 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold106 = new Bold();
            BoldComplexScript boldComplexScript38 = new BoldComplexScript();
            Color color83 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "28" };
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties94.Append(runStyle18);
            runProperties94.Append(runFonts120);
            runProperties94.Append(bold106);
            runProperties94.Append(boldComplexScript38);
            runProperties94.Append(color83);
            runProperties94.Append(fontSizeComplexScript31);
            runProperties94.Append(shading18);
            DeletedText deletedText61 = new DeletedText();
            deletedText61.Text = "[";

            run94.Append(runProperties94);
            run94.Append(deletedText61);

            Run run95 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties95 = new RunProperties();
            RunStyle runStyle19 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts121 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold107 = new Bold();
            BoldComplexScript boldComplexScript39 = new BoldComplexScript();
            Color color84 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "28" };
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages19 = new Languages() { EastAsia = "zh-HK" };

            runProperties95.Append(runStyle19);
            runProperties95.Append(runFonts121);
            runProperties95.Append(bold107);
            runProperties95.Append(boldComplexScript39);
            runProperties95.Append(color84);
            runProperties95.Append(fontSizeComplexScript32);
            runProperties95.Append(shading19);
            runProperties95.Append(languages19);
            DeletedText deletedText62 = new DeletedText();
            deletedText62.Text = "查程工作分配";

            run95.Append(runProperties95);
            run95.Append(deletedText62);

            Run run96 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties96 = new RunProperties();
            RunStyle runStyle20 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts122 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold108 = new Bold();
            BoldComplexScript boldComplexScript40 = new BoldComplexScript();
            Color color85 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "28" };
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties96.Append(runStyle20);
            runProperties96.Append(runFonts122);
            runProperties96.Append(bold108);
            runProperties96.Append(boldComplexScript40);
            runProperties96.Append(color85);
            runProperties96.Append(fontSizeComplexScript33);
            runProperties96.Append(shading20);
            DeletedText deletedText63 = new DeletedText();
            deletedText63.Text = "].[";

            run96.Append(runProperties96);
            run96.Append(deletedText63);

            deletedRun35.Append(run94);
            deletedRun35.Append(run95);
            deletedRun35.Append(run96);

            Run run97 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties97 = new RunProperties();
            RunStyle runStyle21 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts123 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold109 = new Bold();
            BoldComplexScript boldComplexScript41 = new BoldComplexScript();
            Color color86 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "28" };
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages20 = new Languages() { EastAsia = "zh-HK" };

            runProperties97.Append(runStyle21);
            runProperties97.Append(runFonts123);
            runProperties97.Append(bold109);
            runProperties97.Append(boldComplexScript41);
            runProperties97.Append(color86);
            runProperties97.Append(fontSizeComplexScript34);
            runProperties97.Append(shading21);
            runProperties97.Append(languages20);
            Text text34 = new Text();
            text34.Text = "subname";

            run97.Append(runProperties97);
            run97.Append(text34);

            DeletedRun deletedRun36 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "35" };

            Run run98 = new Run() { RsidRunProperties = "00917606", RsidRunDeletion = "0096615A", RsidRunAddition = "00F6052A" };

            RunProperties runProperties98 = new RunProperties();
            RunStyle runStyle22 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts124 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold110 = new Bold();
            BoldComplexScript boldComplexScript42 = new BoldComplexScript();
            Color color87 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "28" };
            Shading shading22 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties98.Append(runStyle22);
            runProperties98.Append(runFonts124);
            runProperties98.Append(bold110);
            runProperties98.Append(boldComplexScript42);
            runProperties98.Append(color87);
            runProperties98.Append(fontSizeComplexScript35);
            runProperties98.Append(shading22);
            DeletedText deletedText64 = new DeletedText();
            deletedText64.Text = "_Eng";

            run98.Append(runProperties98);
            run98.Append(deletedText64);

            deletedRun36.Append(run98);

            DeletedRun deletedRun37 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "36" };

            Run run99 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties99 = new RunProperties();
            RunStyle runStyle23 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts125 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold111 = new Bold();
            BoldComplexScript boldComplexScript43 = new BoldComplexScript();
            Color color88 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "28" };
            Shading shading23 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties99.Append(runStyle23);
            runProperties99.Append(runFonts125);
            runProperties99.Append(bold111);
            runProperties99.Append(boldComplexScript43);
            runProperties99.Append(color88);
            runProperties99.Append(fontSizeComplexScript36);
            runProperties99.Append(shading23);
            DeletedText deletedText65 = new DeletedText();
            deletedText65.Text = "]";

            run99.Append(runProperties99);
            run99.Append(deletedText65);

            deletedRun37.Append(run99);

            paragraph26.Append(paragraphProperties26);
            paragraph26.Append(deletedRun35);
            paragraph26.Append(run97);
            paragraph26.Append(deletedRun36);
            paragraph26.Append(deletedRun37);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph26);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellVerticalAlignment6);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00EB5D9F", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "32A329C8", TextId = "311C84BB" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            SnapToGrid snapToGrid24 = new SnapToGrid() { Val = false };
            Justification justification25 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            RunFonts runFonts126 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties27.Append(runFonts126);

            paragraphProperties27.Append(snapToGrid24);
            paragraphProperties27.Append(justification25);
            paragraphProperties27.Append(paragraphMarkRunProperties27);

            DeletedRun deletedRun38 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "37" };

            Run run100 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties100 = new RunProperties();
            RunFonts runFonts127 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold112 = new Bold();
            Color color89 = new Color() { Val = "0000FF" };

            runProperties100.Append(runFonts127);
            runProperties100.Append(bold112);
            runProperties100.Append(color89);
            DeletedText deletedText66 = new DeletedText();
            deletedText66.Text = "[";

            run100.Append(runProperties100);
            run100.Append(deletedText66);

            Run run101 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties101 = new RunProperties();
            RunFonts runFonts128 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold113 = new Bold();
            Color color90 = new Color() { Val = "0000FF" };

            runProperties101.Append(runFonts128);
            runProperties101.Append(bold113);
            runProperties101.Append(color90);
            DeletedText deletedText67 = new DeletedText();
            deletedText67.Text = "查程工作分配";

            run101.Append(runProperties101);
            run101.Append(deletedText67);

            Run run102 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties102 = new RunProperties();
            RunFonts runFonts129 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold114 = new Bold();
            Color color91 = new Color() { Val = "0000FF" };

            runProperties102.Append(runFonts129);
            runProperties102.Append(bold114);
            runProperties102.Append(color91);
            DeletedText deletedText68 = new DeletedText();
            deletedText68.Text = "].[";

            run102.Append(runProperties102);
            run102.Append(deletedText68);

            deletedRun38.Append(run100);
            deletedRun38.Append(run101);
            deletedRun38.Append(run102);

            Run run103 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties103 = new RunProperties();
            RunFonts runFonts130 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold115 = new Bold();
            Color color92 = new Color() { Val = "0000FF" };

            runProperties103.Append(runFonts130);
            runProperties103.Append(bold115);
            runProperties103.Append(color92);
            Text text35 = new Text();
            text35.Text = dt.Rows[0]["Member"].ToString();

            run103.Append(runProperties103);
            run103.Append(text35);

            DeletedRun deletedRun39 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "38" };

            Run run104 = new Run() { RsidRunProperties = "00917606", RsidRunDeletion = "0096615A", RsidRunAddition = "00F6052A" };

            RunProperties runProperties104 = new RunProperties();
            RunStyle runStyle24 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts131 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold116 = new Bold();
            BoldComplexScript boldComplexScript44 = new BoldComplexScript();
            Color color93 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "28" };
            Shading shading24 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties104.Append(runStyle24);
            runProperties104.Append(runFonts131);
            runProperties104.Append(bold116);
            runProperties104.Append(boldComplexScript44);
            runProperties104.Append(color93);
            runProperties104.Append(fontSizeComplexScript37);
            runProperties104.Append(shading24);
            DeletedText deletedText69 = new DeletedText();
            deletedText69.Text = "_Eng";

            run104.Append(runProperties104);
            run104.Append(deletedText69);

            deletedRun39.Append(run104);

            DeletedRun deletedRun40 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "39" };

            Run run105 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties105 = new RunProperties();
            RunFonts runFonts132 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold117 = new Bold();
            Color color94 = new Color() { Val = "0000FF" };

            runProperties105.Append(runFonts132);
            runProperties105.Append(bold117);
            runProperties105.Append(color94);
            DeletedText deletedText70 = new DeletedText();
            deletedText70.Text = "]";

            run105.Append(runProperties105);
            run105.Append(deletedText70);

            deletedRun40.Append(run105);

            paragraph27.Append(paragraphProperties27);
            paragraph27.Append(deletedRun38);
            paragraph27.Append(run103);
            paragraph27.Append(deletedRun39);
            paragraph27.Append(deletedRun40);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph27);

            tableRow10.Append(tableRowProperties1);
            tableRow10.Append(tableCell21);
            tableRow10.Append(tableCell22);
            tableRow10.Append(tableCell23);

            TableRow tableRow11 = new TableRow() { RsidTableRowMarkRevision = "0096027D", RsidTableRowAddition = "00EE6462", RsidTableRowDeletion = "001C7370", RsidTableRowProperties = "00193FEF", ParagraphId = "76DCE3E2", TextId = "1166F2DA" };

            TableRowProperties tableRowProperties2 = new TableRowProperties();
            TableRowHeight tableRowHeight2 = new TableRowHeight() { Val = (UInt32Value)80U };
            Deleted deleted1 = new Deleted() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "40" };

            tableRowProperties2.Append(tableRowHeight2);
            tableRowProperties2.Append(deleted1);

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment7 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties24.Append(tableCellWidth24);
            tableCellProperties24.Append(tableCellVerticalAlignment7);

            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00EE6462", RsidParagraphDeletion = "001C7370", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EE6462", ParagraphId = "1FB69D4A", TextId = "777BD2CD" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            SnapToGrid snapToGrid25 = new SnapToGrid() { Val = false };
            Justification justification26 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            Deleted deleted2 = new Deleted() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "41" };
            RunStyle runStyle25 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts133 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold118 = new Bold();
            BoldComplexScript boldComplexScript45 = new BoldComplexScript();
            Color color95 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "28" };
            Shading shading25 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            paragraphMarkRunProperties28.Append(deleted2);
            paragraphMarkRunProperties28.Append(runStyle25);
            paragraphMarkRunProperties28.Append(runFonts133);
            paragraphMarkRunProperties28.Append(bold118);
            paragraphMarkRunProperties28.Append(boldComplexScript45);
            paragraphMarkRunProperties28.Append(color95);
            paragraphMarkRunProperties28.Append(fontSizeComplexScript38);
            paragraphMarkRunProperties28.Append(shading25);

            paragraphProperties28.Append(snapToGrid25);
            paragraphProperties28.Append(justification26);
            paragraphProperties28.Append(paragraphMarkRunProperties28);

            DeletedRun deletedRun41 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "42" };

            Run run106 = new Run() { RsidRunProperties = "0096027D", RsidRunDeletion = "001C7370" };

            RunProperties runProperties106 = new RunProperties();
            RunStyle runStyle26 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts134 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold119 = new Bold();
            BoldComplexScript boldComplexScript46 = new BoldComplexScript();
            Color color96 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "28" };
            Shading shading26 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            runProperties106.Append(runStyle26);
            runProperties106.Append(runFonts134);
            runProperties106.Append(bold119);
            runProperties106.Append(boldComplexScript46);
            runProperties106.Append(color96);
            runProperties106.Append(fontSizeComplexScript39);
            runProperties106.Append(shading26);
            DeletedText deletedText71 = new DeletedText();
            deletedText71.Text = "….";

            run106.Append(runProperties106);
            run106.Append(deletedText71);

            deletedRun41.Append(run106);

            paragraph28.Append(paragraphProperties28);
            paragraph28.Append(deletedRun41);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph28);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment8 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties25.Append(tableCellWidth25);
            tableCellProperties25.Append(tableCellVerticalAlignment8);

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00EE6462", RsidParagraphDeletion = "001C7370", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EE6462", ParagraphId = "3BBCEAD0", TextId = "6452ED3E" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();
            SnapToGrid snapToGrid26 = new SnapToGrid() { Val = false };
            Indentation indentation18 = new Indentation() { FirstLine = "240", FirstLineChars = 100 };
            Justification justification27 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            Deleted deleted3 = new Deleted() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "43" };
            RunStyle runStyle27 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts135 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold120 = new Bold();
            BoldComplexScript boldComplexScript47 = new BoldComplexScript();
            Color color97 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "28" };
            Shading shading27 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };

            paragraphMarkRunProperties29.Append(deleted3);
            paragraphMarkRunProperties29.Append(runStyle27);
            paragraphMarkRunProperties29.Append(runFonts135);
            paragraphMarkRunProperties29.Append(bold120);
            paragraphMarkRunProperties29.Append(boldComplexScript47);
            paragraphMarkRunProperties29.Append(color97);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript40);
            paragraphMarkRunProperties29.Append(shading27);

            paragraphProperties29.Append(snapToGrid26);
            paragraphProperties29.Append(indentation18);
            paragraphProperties29.Append(justification27);
            paragraphProperties29.Append(paragraphMarkRunProperties29);

            paragraph29.Append(paragraphProperties29);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph29);

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment9 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties26.Append(tableCellWidth26);
            tableCellProperties26.Append(tableCellVerticalAlignment9);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00EE6462", RsidParagraphDeletion = "001C7370", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EE6462", ParagraphId = "5988F99D", TextId = "5C645D95" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();
            SnapToGrid snapToGrid27 = new SnapToGrid() { Val = false };
            Justification justification28 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            Deleted deleted4 = new Deleted() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "44" };
            RunFonts runFonts136 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold121 = new Bold();
            Color color98 = new Color() { Val = "0000FF" };

            paragraphMarkRunProperties30.Append(deleted4);
            paragraphMarkRunProperties30.Append(runFonts136);
            paragraphMarkRunProperties30.Append(bold121);
            paragraphMarkRunProperties30.Append(color98);

            paragraphProperties30.Append(snapToGrid27);
            paragraphProperties30.Append(justification28);
            paragraphProperties30.Append(paragraphMarkRunProperties30);

            paragraph30.Append(paragraphProperties30);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph30);

            tableRow11.Append(tableRowProperties2);
            tableRow11.Append(tableCell24);
            tableRow11.Append(tableCell25);
            tableRow11.Append(tableCell26);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow8);
            table2.Append(tableRow9);
            table2.Append(tableRow10);
            table2.Append(tableRow11);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00405678", RsidParagraphProperties = "000E3880", RsidRunAdditionDefault = "00405678", ParagraphId = "131F4BDE", TextId = "77777777" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            RunFonts runFonts137 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties31.Append(runFonts137);

            paragraphProperties31.Append(paragraphMarkRunProperties31);

            paragraph31.Append(paragraphProperties31);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00715EA3", RsidRunAdditionDefault = "00715EA3", ParagraphId = "32AE3361", TextId = "77777777" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();

            NumberingProperties numberingProperties2 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference2 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId2 = new NumberingId() { Val = 15 };

            numberingProperties2.Append(numberingLevelReference2);
            numberingProperties2.Append(numberingId2);
            SnapToGrid snapToGrid28 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            RunFonts runFonts138 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold122 = new Bold();
            BoldComplexScript boldComplexScript48 = new BoldComplexScript();
            FontSize fontSize92 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties32.Append(runFonts138);
            paragraphMarkRunProperties32.Append(bold122);
            paragraphMarkRunProperties32.Append(boldComplexScript48);
            paragraphMarkRunProperties32.Append(fontSize92);

            paragraphProperties32.Append(numberingProperties2);
            paragraphProperties32.Append(snapToGrid28);
            paragraphProperties32.Append(paragraphMarkRunProperties32);

            Run run107 = new Run() { RsidRunProperties = "00715EA3" };

            RunProperties runProperties107 = new RunProperties();
            RunFonts runFonts139 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold123 = new Bold();
            BoldComplexScript boldComplexScript49 = new BoldComplexScript();
            FontSize fontSize93 = new FontSize() { Val = "28" };

            runProperties107.Append(runFonts139);
            runProperties107.Append(bold123);
            runProperties107.Append(boldComplexScript49);
            runProperties107.Append(fontSize93);
            Text text36 = new Text();
            text36.Text = "Overview information";

            run107.Append(runProperties107);
            run107.Append(text36);

            Run run108 = new Run() { RsidRunAddition = "007F22B3" };

            RunProperties runProperties108 = new RunProperties();
            RunFonts runFonts140 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold124 = new Bold();
            BoldComplexScript boldComplexScript50 = new BoldComplexScript();
            FontSize fontSize94 = new FontSize() { Val = "28" };

            runProperties108.Append(runFonts140);
            runProperties108.Append(bold124);
            runProperties108.Append(boldComplexScript50);
            runProperties108.Append(fontSize94);
            Text text37 = new Text();
            text37.Text = ":";

            run108.Append(runProperties108);
            run108.Append(text37);

            paragraph32.Append(paragraphProperties32);
            paragraph32.Append(run107);
            paragraph32.Append(run108);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "007F22B3", RsidParagraphAddition = "00715EA3", RsidParagraphProperties = "00715EA3", RsidRunAdditionDefault = "00715EA3", ParagraphId = "5FF5DEC3", TextId = "77777777" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            SnapToGrid snapToGrid29 = new SnapToGrid() { Val = false };
            Indentation indentation19 = new Indentation() { Start = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            RunFonts runFonts141 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold125 = new Bold();
            BoldComplexScript boldComplexScript51 = new BoldComplexScript();
            FontSize fontSize95 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties33.Append(runFonts141);
            paragraphMarkRunProperties33.Append(bold125);
            paragraphMarkRunProperties33.Append(boldComplexScript51);
            paragraphMarkRunProperties33.Append(fontSize95);

            paragraphProperties33.Append(snapToGrid29);
            paragraphProperties33.Append(indentation19);
            paragraphProperties33.Append(paragraphMarkRunProperties33);

            paragraph33.Append(paragraphProperties33);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "00715EA3", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00715EA3", RsidRunAdditionDefault = "00715EA3", ParagraphId = "766FC676", TextId = "77777777" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();

            NumberingProperties numberingProperties3 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference3 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId3 = new NumberingId() { Val = 15 };

            numberingProperties3.Append(numberingLevelReference3);
            numberingProperties3.Append(numberingId3);
            SnapToGrid snapToGrid30 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts142 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold126 = new Bold();
            BoldComplexScript boldComplexScript52 = new BoldComplexScript();
            FontSize fontSize96 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties34.Append(runFonts142);
            paragraphMarkRunProperties34.Append(bold126);
            paragraphMarkRunProperties34.Append(boldComplexScript52);
            paragraphMarkRunProperties34.Append(fontSize96);

            paragraphProperties34.Append(numberingProperties3);
            paragraphProperties34.Append(snapToGrid30);
            paragraphProperties34.Append(paragraphMarkRunProperties34);

            Run run109 = new Run();

            RunProperties runProperties109 = new RunProperties();
            RunFonts runFonts143 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold127 = new Bold();
            BoldComplexScript boldComplexScript53 = new BoldComplexScript();
            FontSize fontSize97 = new FontSize() { Val = "28" };

            runProperties109.Append(runFonts143);
            runProperties109.Append(bold127);
            runProperties109.Append(boldComplexScript53);
            runProperties109.Append(fontSize97);
            Text text38 = new Text();
            text38.Text = "M";

            run109.Append(runProperties109);
            run109.Append(text38);

            Run run110 = new Run() { RsidRunAddition = "00C34568" };

            RunProperties runProperties110 = new RunProperties();
            RunFonts runFonts144 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold128 = new Bold();
            BoldComplexScript boldComplexScript54 = new BoldComplexScript();
            FontSize fontSize98 = new FontSize() { Val = "28" };

            runProperties110.Append(runFonts144);
            runProperties110.Append(bold128);
            runProperties110.Append(boldComplexScript54);
            runProperties110.Append(fontSize98);
            Text text39 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text39.Text = "aterial changes ";

            run110.Append(runProperties110);
            run110.Append(text39);

            Run run111 = new Run() { RsidRunAddition = "00C34568" };

            RunProperties runProperties111 = new RunProperties();
            RunFonts runFonts145 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold129 = new Bold();
            BoldComplexScript boldComplexScript55 = new BoldComplexScript();
            FontSize fontSize99 = new FontSize() { Val = "28" };

            runProperties111.Append(runFonts145);
            runProperties111.Append(bold129);
            runProperties111.Append(boldComplexScript55);
            runProperties111.Append(fontSize99);
            Text text40 = new Text();
            text40.Text = "in the:";

            run111.Append(runProperties111);
            run111.Append(text40);

            paragraph34.Append(paragraphProperties34);
            paragraph34.Append(run109);
            paragraph34.Append(run110);
            paragraph34.Append(run111);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "008C261F", RsidRunAdditionDefault = "00A244BE", ParagraphId = "391DA516", TextId = "77777777" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();

            NumberingProperties numberingProperties4 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference4 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId4 = new NumberingId() { Val = 16 };

            numberingProperties4.Append(numberingLevelReference4);
            numberingProperties4.Append(numberingId4);

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Number, Position = 1134 };

            tabs1.Append(tabStop1);
            SnapToGrid snapToGrid31 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts146 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize100 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties35.Append(runFonts146);
            paragraphMarkRunProperties35.Append(fontSize100);

            paragraphProperties35.Append(numberingProperties4);
            paragraphProperties35.Append(tabs1);
            paragraphProperties35.Append(snapToGrid31);
            paragraphProperties35.Append(spacingBetweenLines1);
            paragraphProperties35.Append(paragraphMarkRunProperties35);

            Run run112 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties112 = new RunProperties();
            RunFonts runFonts147 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize101 = new FontSize() { Val = "28" };

            runProperties112.Append(runFonts147);
            runProperties112.Append(fontSize101);
            Text text41 = new Text();
            text41.Text = "Organization:";

            run112.Append(runProperties112);
            run112.Append(text41);

            paragraph35.Append(paragraphProperties35);
            paragraph35.Append(run112);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00DD4C3C", RsidParagraphProperties = "008C261F", RsidRunAdditionDefault = "009A6A85", ParagraphId = "41531F86", TextId = "77777777" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();

            NumberingProperties numberingProperties5 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference5 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId5 = new NumberingId() { Val = 16 };

            numberingProperties5.Append(numberingLevelReference5);
            numberingProperties5.Append(numberingId5);

            Tabs tabs2 = new Tabs();
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Number, Position = 1134 };

            tabs2.Append(tabStop2);
            SnapToGrid snapToGrid32 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts148 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize102 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties36.Append(runFonts148);
            paragraphMarkRunProperties36.Append(fontSize102);

            paragraphProperties36.Append(numberingProperties5);
            paragraphProperties36.Append(tabs2);
            paragraphProperties36.Append(snapToGrid32);
            paragraphProperties36.Append(spacingBetweenLines2);
            paragraphProperties36.Append(paragraphMarkRunProperties36);

            Run run113 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties113 = new RunProperties();
            RunFonts runFonts149 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize103 = new FontSize() { Val = "28" };

            runProperties113.Append(runFonts149);
            runProperties113.Append(fontSize103);
            Text text42 = new Text();
            text42.Text = "Bu";

            run113.Append(runProperties113);
            run113.Append(text42);

            Run run114 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties114 = new RunProperties();
            RunFonts runFonts150 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize104 = new FontSize() { Val = "28" };

            runProperties114.Append(runFonts150);
            runProperties114.Append(fontSize104);
            Text text43 = new Text();
            text43.Text = "siness & Operation";

            run114.Append(runProperties114);
            run114.Append(text43);

            Run run115 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "00DD4C3C" };

            RunProperties runProperties115 = new RunProperties();
            RunFonts runFonts151 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize105 = new FontSize() { Val = "28" };

            runProperties115.Append(runFonts151);
            runProperties115.Append(fontSize105);
            Text text44 = new Text();
            text44.Text = ":";

            run115.Append(runProperties115);
            run115.Append(text44);

            paragraph36.Append(paragraphProperties36);
            paragraph36.Append(run113);
            paragraph36.Append(run114);
            paragraph36.Append(run115);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "008C261F", RsidRunAdditionDefault = "000D4BC6", ParagraphId = "2CED2C0D", TextId = "77777777" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();

            NumberingProperties numberingProperties6 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference6 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId6 = new NumberingId() { Val = 16 };

            numberingProperties6.Append(numberingLevelReference6);
            numberingProperties6.Append(numberingId6);

            Tabs tabs3 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Number, Position = 1134 };

            tabs3.Append(tabStop3);
            SnapToGrid snapToGrid33 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts152 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize106 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties37.Append(runFonts152);
            paragraphMarkRunProperties37.Append(fontSize106);

            paragraphProperties37.Append(numberingProperties6);
            paragraphProperties37.Append(tabs3);
            paragraphProperties37.Append(snapToGrid33);
            paragraphProperties37.Append(spacingBetweenLines3);
            paragraphProperties37.Append(paragraphMarkRunProperties37);

            Run run116 = new Run();

            RunProperties runProperties116 = new RunProperties();
            RunFonts runFonts153 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize107 = new FontSize() { Val = "28" };

            runProperties116.Append(runFonts153);
            runProperties116.Append(fontSize107);
            Text text45 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text45.Text = "Key ";

            run116.Append(runProperties116);
            run116.Append(text45);

            Run run117 = new Run() { RsidRunProperties = "00DD4C3C", RsidRunAddition = "00DD4C3C" };

            RunProperties runProperties117 = new RunProperties();
            RunFonts runFonts154 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize108 = new FontSize() { Val = "28" };

            runProperties117.Append(runFonts154);
            runProperties117.Append(fontSize108);
            Text text46 = new Text();
            text46.Text = "Regulat";

            run117.Append(runProperties117);
            run117.Append(text46);

            Run run118 = new Run() { RsidRunAddition = "003F3517" };

            RunProperties runProperties118 = new RunProperties();
            RunFonts runFonts155 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize109 = new FontSize() { Val = "28" };

            runProperties118.Append(runFonts155);
            runProperties118.Append(fontSize109);
            Text text47 = new Text();
            text47.Text = "io";

            run118.Append(runProperties118);
            run118.Append(text47);

            Run run119 = new Run() { RsidRunAddition = "003F3517" };

            RunProperties runProperties119 = new RunProperties();
            RunFonts runFonts156 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize110 = new FontSize() { Val = "28" };

            runProperties119.Append(runFonts156);
            runProperties119.Append(fontSize110);
            Text text48 = new Text();
            text48.Text = "n";

            run119.Append(runProperties119);
            run119.Append(text48);

            Run run120 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "006C3779" };

            RunProperties runProperties120 = new RunProperties();
            RunStyle runStyle28 = new RunStyle() { Val = "a9" };
            RunFonts runFonts157 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize111 = new FontSize() { Val = "28" };

            runProperties120.Append(runStyle28);
            runProperties120.Append(runFonts157);
            runProperties120.Append(fontSize111);
            FootnoteReference footnoteReference1 = new FootnoteReference() { Id = 1 };

            run120.Append(runProperties120);
            run120.Append(footnoteReference1);

            Run run121 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "00DD4C3C" };

            RunProperties runProperties121 = new RunProperties();
            RunFonts runFonts158 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize112 = new FontSize() { Val = "28" };

            runProperties121.Append(runFonts158);
            runProperties121.Append(fontSize112);
            Text text49 = new Text();
            text49.Text = ":";

            run121.Append(runProperties121);
            run121.Append(text49);

            paragraph37.Append(paragraphProperties37);
            paragraph37.Append(run116);
            paragraph37.Append(run117);
            paragraph37.Append(run118);
            paragraph37.Append(run119);
            paragraph37.Append(run120);
            paragraph37.Append(run121);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00E30AF0", RsidParagraphProperties = "00E30AF0", RsidRunAdditionDefault = "00E30AF0", ParagraphId = "21BFB66E", TextId = "77777777" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Number, Position = 1680 };

            tabs4.Append(tabStop4);
            SnapToGrid snapToGrid34 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation20 = new Indentation() { Start = "960" };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts159 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize113 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties38.Append(runFonts159);
            paragraphMarkRunProperties38.Append(fontSize113);

            paragraphProperties38.Append(tabs4);
            paragraphProperties38.Append(snapToGrid34);
            paragraphProperties38.Append(spacingBetweenLines4);
            paragraphProperties38.Append(indentation20);
            paragraphProperties38.Append(paragraphMarkRunProperties38);

            paragraph38.Append(paragraphProperties38);

            Paragraph paragraph39 = new Paragraph() { RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00FC58E2", RsidRunAdditionDefault = "00E30AF0", ParagraphId = "52C66D48", TextId = "1E393CF8" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();

            NumberingProperties numberingProperties7 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference7 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId7 = new NumberingId() { Val = 15 };

            numberingProperties7.Append(numberingLevelReference7);
            numberingProperties7.Append(numberingId7);
            SnapToGrid snapToGrid35 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts160 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold130 = new Bold();
            BoldComplexScript boldComplexScript56 = new BoldComplexScript();
            FontSize fontSize114 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties39.Append(runFonts160);
            paragraphMarkRunProperties39.Append(bold130);
            paragraphMarkRunProperties39.Append(boldComplexScript56);
            paragraphMarkRunProperties39.Append(fontSize114);

            paragraphProperties39.Append(numberingProperties7);
            paragraphProperties39.Append(snapToGrid35);
            paragraphProperties39.Append(paragraphMarkRunProperties39);

            Run run122 = new Run() { RsidRunProperties = "00E30AF0" };

            RunProperties runProperties122 = new RunProperties();
            RunFonts runFonts161 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold131 = new Bold();
            BoldComplexScript boldComplexScript57 = new BoldComplexScript();
            FontSize fontSize115 = new FontSize() { Val = "28" };

            runProperties122.Append(runFonts161);
            runProperties122.Append(bold131);
            runProperties122.Append(boldComplexScript57);
            runProperties122.Append(fontSize115);
            Text text50 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text50.Text = "Prior ";

            run122.Append(runProperties122);
            run122.Append(text50);

            Run run123 = new Run() { RsidRunAddition = "00D66E75" };

            RunProperties runProperties123 = new RunProperties();
            RunFonts runFonts162 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold132 = new Bold();
            BoldComplexScript boldComplexScript58 = new BoldComplexScript();
            FontSize fontSize116 = new FontSize() { Val = "28" };

            runProperties123.Append(runFonts162);
            runProperties123.Append(bold132);
            runProperties123.Append(boldComplexScript58);
            runProperties123.Append(fontSize116);
            Text text51 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text51.Text = "major ";

            run123.Append(runProperties123);
            run123.Append(text51);

            Run run124 = new Run() { RsidRunProperties = "00E30AF0" };

            RunProperties runProperties124 = new RunProperties();
            RunFonts runFonts163 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold133 = new Bold();
            BoldComplexScript boldComplexScript59 = new BoldComplexScript();
            FontSize fontSize117 = new FontSize() { Val = "28" };

            runProperties124.Append(runFonts163);
            runProperties124.Append(bold133);
            runProperties124.Append(boldComplexScript59);
            runProperties124.Append(fontSize117);
            Text text52 = new Text();
            text52.Text = "audit issues";

            run124.Append(runProperties124);
            run124.Append(text52);

            Run run125 = new Run();

            RunProperties runProperties125 = new RunProperties();
            RunFonts runFonts164 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold134 = new Bold();
            BoldComplexScript boldComplexScript60 = new BoldComplexScript();
            FontSize fontSize118 = new FontSize() { Val = "28" };

            runProperties125.Append(runFonts164);
            runProperties125.Append(bold134);
            runProperties125.Append(boldComplexScript60);
            runProperties125.Append(fontSize118);
            Text text53 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text53.Text = " ";

            run125.Append(runProperties125);
            run125.Append(text53);

            Run run126 = new Run() { RsidRunAddition = "00FC58E2" };

            RunProperties runProperties126 = new RunProperties();
            RunFonts runFonts165 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold135 = new Bold();
            BoldComplexScript boldComplexScript61 = new BoldComplexScript();
            FontSize fontSize119 = new FontSize() { Val = "28" };

            runProperties126.Append(runFonts165);
            runProperties126.Append(bold135);
            runProperties126.Append(boldComplexScript61);
            runProperties126.Append(fontSize119);
            Text text54 = new Text();
            text54.Text = "(";

            run126.Append(runProperties126);
            run126.Append(text54);

            Run run127 = new Run() { RsidRunAddition = "00FD24E6" };

            RunProperties runProperties127 = new RunProperties();
            RunFonts runFonts166 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold136 = new Bold();
            BoldComplexScript boldComplexScript62 = new BoldComplexScript();
            FontSize fontSize120 = new FontSize() { Val = "28" };

            runProperties127.Append(runFonts166);
            runProperties127.Append(bold136);
            runProperties127.Append(boldComplexScript62);
            runProperties127.Append(fontSize120);
            Text text55 = new Text();
            text55.Text = "include:";

            run127.Append(runProperties127);
            run127.Append(text55);

            Run run128 = new Run() { RsidRunAddition = "00D66E75" };

            RunProperties runProperties128 = new RunProperties();
            RunFonts runFonts167 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold137 = new Bold();
            BoldComplexScript boldComplexScript63 = new BoldComplexScript();
            FontSize fontSize121 = new FontSize() { Val = "28" };

            runProperties128.Append(runFonts167);
            runProperties128.Append(bold137);
            runProperties128.Append(boldComplexScript63);
            runProperties128.Append(fontSize121);
            Text text56 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text56.Text = " ";

            run128.Append(runProperties128);
            run128.Append(text56);

            Run run129 = new Run() { RsidRunAddition = "00FC58E2" };

            RunProperties runProperties129 = new RunProperties();
            RunFonts runFonts168 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold138 = new Bold();
            BoldComplexScript boldComplexScript64 = new BoldComplexScript();
            FontSize fontSize122 = new FontSize() { Val = "28" };

            runProperties129.Append(runFonts168);
            runProperties129.Append(bold138);
            runProperties129.Append(boldComplexScript64);
            runProperties129.Append(fontSize122);
            Text text57 = new Text();
            text57.Text = "r";

            run129.Append(runProperties129);
            run129.Append(text57);

            Run run130 = new Run() { RsidRunProperties = "00FC58E2", RsidRunAddition = "00FC58E2" };

            RunProperties runProperties130 = new RunProperties();
            RunFonts runFonts169 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold139 = new Bold();
            BoldComplexScript boldComplexScript65 = new BoldComplexScript();
            FontSize fontSize123 = new FontSize() { Val = "28" };

            runProperties130.Append(runFonts169);
            runProperties130.Append(bold139);
            runProperties130.Append(boldComplexScript65);
            runProperties130.Append(fontSize123);
            Text text58 = new Text();
            text58.Text = "egulator";

            run130.Append(runProperties130);
            run130.Append(text58);

            Run run131 = new Run() { RsidRunAddition = "00FC58E2" };

            RunProperties runProperties131 = new RunProperties();
            RunFonts runFonts170 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold140 = new Bold();
            BoldComplexScript boldComplexScript66 = new BoldComplexScript();
            FontSize fontSize124 = new FontSize() { Val = "28" };

            runProperties131.Append(runFonts170);
            runProperties131.Append(bold140);
            runProperties131.Append(boldComplexScript66);
            runProperties131.Append(fontSize124);
            Text text59 = new Text();
            text59.Text = ",";

            run131.Append(runProperties131);
            run131.Append(text59);

            Run run132 = new Run() { RsidRunProperties = "00FC58E2", RsidRunAddition = "00FC58E2" };

            RunProperties runProperties132 = new RunProperties();
            RunFonts runFonts171 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold141 = new Bold();
            BoldComplexScript boldComplexScript67 = new BoldComplexScript();
            FontSize fontSize125 = new FontSize() { Val = "28" };

            runProperties132.Append(runFonts171);
            runProperties132.Append(bold141);
            runProperties132.Append(boldComplexScript67);
            runProperties132.Append(fontSize125);
            Text text60 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text60.Text = " ";

            run132.Append(runProperties132);
            run132.Append(text60);

            Run run133 = new Run();

            RunProperties runProperties133 = new RunProperties();
            RunFonts runFonts172 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold142 = new Bold();
            BoldComplexScript boldComplexScript68 = new BoldComplexScript();
            FontSize fontSize126 = new FontSize() { Val = "28" };

            runProperties133.Append(runFonts172);
            runProperties133.Append(bold142);
            runProperties133.Append(boldComplexScript68);
            runProperties133.Append(fontSize126);
            Text text61 = new Text();
            text61.Text = "internal";

            run133.Append(runProperties133);
            run133.Append(text61);

            Run run134 = new Run() { RsidRunAddition = "00FC58E2" };

            RunProperties runProperties134 = new RunProperties();
            RunFonts runFonts173 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold143 = new Bold();
            BoldComplexScript boldComplexScript69 = new BoldComplexScript();
            FontSize fontSize127 = new FontSize() { Val = "28" };

            runProperties134.Append(runFonts173);
            runProperties134.Append(bold143);
            runProperties134.Append(boldComplexScript69);
            runProperties134.Append(fontSize127);
            Text text62 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text62.Text = " audit,";

            run134.Append(runProperties134);
            run134.Append(text62);

            Run run135 = new Run();

            RunProperties runProperties135 = new RunProperties();
            RunFonts runFonts174 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold144 = new Bold();
            BoldComplexScript boldComplexScript70 = new BoldComplexScript();
            FontSize fontSize128 = new FontSize() { Val = "28" };

            runProperties135.Append(runFonts174);
            runProperties135.Append(bold144);
            runProperties135.Append(boldComplexScript70);
            runProperties135.Append(fontSize128);
            Text text63 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text63.Text = " external";

            run135.Append(runProperties135);
            run135.Append(text63);

            Run run136 = new Run() { RsidRunAddition = "00FC58E2" };

            RunProperties runProperties136 = new RunProperties();
            RunFonts runFonts175 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold145 = new Bold();
            BoldComplexScript boldComplexScript71 = new BoldComplexScript();
            FontSize fontSize129 = new FontSize() { Val = "28" };

            runProperties136.Append(runFonts175);
            runProperties136.Append(bold145);
            runProperties136.Append(boldComplexScript71);
            runProperties136.Append(fontSize129);
            Text text64 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text64.Text = " audit";

            run136.Append(runProperties136);
            run136.Append(text64);

            Run run137 = new Run();

            RunProperties runProperties137 = new RunProperties();
            RunFonts runFonts176 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold146 = new Bold();
            BoldComplexScript boldComplexScript72 = new BoldComplexScript();
            FontSize fontSize130 = new FontSize() { Val = "28" };

            runProperties137.Append(runFonts176);
            runProperties137.Append(bold146);
            runProperties137.Append(boldComplexScript72);
            runProperties137.Append(fontSize130);
            Text text65 = new Text();
            text65.Text = ")";

            run137.Append(runProperties137);
            run137.Append(text65);

            Run run138 = new Run() { RsidRunProperties = "00231BC8", RsidRunAddition = "00CF43E8" };

            RunProperties runProperties138 = new RunProperties();
            RunFonts runFonts177 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold147 = new Bold();
            BoldComplexScript boldComplexScript73 = new BoldComplexScript();
            FontSize fontSize131 = new FontSize() { Val = "28" };
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties138.Append(runFonts177);
            runProperties138.Append(bold147);
            runProperties138.Append(boldComplexScript73);
            runProperties138.Append(fontSize131);
            runProperties138.Append(verticalTextAlignment1);
            FootnoteReference footnoteReference2 = new FootnoteReference() { Id = 2 };

            run138.Append(runProperties138);
            run138.Append(footnoteReference2);

            Run run139 = new Run() { RsidRunProperties = "00E30AF0", RsidRunAddition = "0014137D" };

            RunProperties runProperties139 = new RunProperties();
            RunFonts runFonts178 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold148 = new Bold();
            BoldComplexScript boldComplexScript74 = new BoldComplexScript();
            FontSize fontSize132 = new FontSize() { Val = "28" };

            runProperties139.Append(runFonts178);
            runProperties139.Append(bold148);
            runProperties139.Append(boldComplexScript74);
            runProperties139.Append(fontSize132);
            Text text66 = new Text();
            text66.Text = "：";

            run139.Append(runProperties139);
            run139.Append(text66);

            paragraph39.Append(paragraphProperties39);
            paragraph39.Append(run122);
            paragraph39.Append(run123);
            paragraph39.Append(run124);
            paragraph39.Append(run125);
            paragraph39.Append(run126);
            paragraph39.Append(run127);
            paragraph39.Append(run128);
            paragraph39.Append(run129);
            paragraph39.Append(run130);
            paragraph39.Append(run131);
            paragraph39.Append(run132);
            paragraph39.Append(run133);
            paragraph39.Append(run134);
            paragraph39.Append(run135);
            paragraph39.Append(run136);
            paragraph39.Append(run137);
            paragraph39.Append(run138);
            paragraph39.Append(run139);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphMarkRevision = "00FD24E6", RsidParagraphAddition = "00D02A33", RsidParagraphProperties = "00D02A33", RsidRunAdditionDefault = "00D02A33", ParagraphId = "215B3C66", TextId = "77777777" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();
            SnapToGrid snapToGrid36 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts runFonts179 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold149 = new Bold();
            BoldComplexScript boldComplexScript75 = new BoldComplexScript();
            FontSize fontSize133 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties40.Append(runFonts179);
            paragraphMarkRunProperties40.Append(bold149);
            paragraphMarkRunProperties40.Append(boldComplexScript75);
            paragraphMarkRunProperties40.Append(fontSize133);

            paragraphProperties40.Append(snapToGrid36);
            paragraphProperties40.Append(paragraphMarkRunProperties40);

            paragraph40.Append(paragraphProperties40);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "00D02A33", RsidParagraphAddition = "00CF43E8", RsidParagraphProperties = "00324D5C", RsidRunAdditionDefault = "00AE1DA4", ParagraphId = "4428D558", TextId = "56711518" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();

            NumberingProperties numberingProperties8 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference8 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId8 = new NumberingId() { Val = 15 };

            numberingProperties8.Append(numberingLevelReference8);
            numberingProperties8.Append(numberingId8);
            SnapToGrid snapToGrid37 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts runFonts180 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold150 = new Bold();
            BoldComplexScript boldComplexScript76 = new BoldComplexScript();
            FontSize fontSize134 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties41.Append(runFonts180);
            paragraphMarkRunProperties41.Append(bold150);
            paragraphMarkRunProperties41.Append(boldComplexScript76);
            paragraphMarkRunProperties41.Append(fontSize134);

            paragraphProperties41.Append(numberingProperties8);
            paragraphProperties41.Append(snapToGrid37);
            paragraphProperties41.Append(paragraphMarkRunProperties41);

            Run run140 = new Run();

            RunProperties runProperties140 = new RunProperties();
            RunFonts runFonts181 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold151 = new Bold();
            BoldComplexScript boldComplexScript77 = new BoldComplexScript();
            FontSize fontSize135 = new FontSize() { Val = "28" };

            runProperties140.Append(runFonts181);
            runProperties140.Append(bold151);
            runProperties140.Append(boldComplexScript77);
            runProperties140.Append(fontSize135);
            Text text67 = new Text();
            text67.Text = "A";

            run140.Append(runProperties140);
            run140.Append(text67);

            Run run141 = new Run();

            RunProperties runProperties141 = new RunProperties();
            RunFonts runFonts182 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold152 = new Bold();
            BoldComplexScript boldComplexScript78 = new BoldComplexScript();
            FontSize fontSize136 = new FontSize() { Val = "28" };

            runProperties141.Append(runFonts182);
            runProperties141.Append(bold152);
            runProperties141.Append(boldComplexScript78);
            runProperties141.Append(fontSize136);
            Text text68 = new Text();
            text68.Text = "udit focus and sampling";

            run141.Append(runProperties141);
            run141.Append(text68);

            Run run142 = new Run() { RsidRunAddition = "006D6273" };

            RunProperties runProperties142 = new RunProperties();
            RunFonts runFonts183 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold153 = new Bold();
            BoldComplexScript boldComplexScript79 = new BoldComplexScript();
            FontSize fontSize137 = new FontSize() { Val = "28" };

            runProperties142.Append(runFonts183);
            runProperties142.Append(bold153);
            runProperties142.Append(boldComplexScript79);
            runProperties142.Append(fontSize137);
            Text text69 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text69.Text = " (";

            run142.Append(runProperties142);
            run142.Append(text69);

            Run run143 = new Run() { RsidRunAddition = "004C7B6F" };

            RunProperties runProperties143 = new RunProperties();
            RunFonts runFonts184 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold154 = new Bold();
            BoldComplexScript boldComplexScript80 = new BoldComplexScript();
            FontSize fontSize138 = new FontSize() { Val = "28" };

            runProperties143.Append(runFonts184);
            runProperties143.Append(bold154);
            runProperties143.Append(boldComplexScript80);
            runProperties143.Append(fontSize138);
            Text text70 = new Text();
            text70.Text = "includ";

            run143.Append(runProperties143);
            run143.Append(text70);

            Run run144 = new Run() { RsidRunAddition = "00324D5C" };

            RunProperties runProperties144 = new RunProperties();
            RunFonts runFonts185 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold155 = new Bold();
            BoldComplexScript boldComplexScript81 = new BoldComplexScript();
            FontSize fontSize139 = new FontSize() { Val = "28" };

            runProperties144.Append(runFonts185);
            runProperties144.Append(bold155);
            runProperties144.Append(boldComplexScript81);
            runProperties144.Append(fontSize139);
            Text text71 = new Text();
            text71.Text = "e";

            run144.Append(runProperties144);
            run144.Append(text71);

            Run run145 = new Run() { RsidRunAddition = "004C7B6F" };

            RunProperties runProperties145 = new RunProperties();
            RunFonts runFonts186 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold156 = new Bold();
            BoldComplexScript boldComplexScript82 = new BoldComplexScript();
            FontSize fontSize140 = new FontSize() { Val = "28" };

            runProperties145.Append(runFonts186);
            runProperties145.Append(bold156);
            runProperties145.Append(boldComplexScript82);
            runProperties145.Append(fontSize140);
            Text text72 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text72.Text = " but not limited to:";

            run145.Append(runProperties145);
            run145.Append(text72);

            Run run146 = new Run() { RsidRunAddition = "000B4F5B" };

            RunProperties runProperties146 = new RunProperties();
            RunFonts runFonts187 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold157 = new Bold();
            BoldComplexScript boldComplexScript83 = new BoldComplexScript();
            FontSize fontSize141 = new FontSize() { Val = "28" };

            runProperties146.Append(runFonts187);
            runProperties146.Append(bold157);
            runProperties146.Append(boldComplexScript83);
            runProperties146.Append(fontSize141);
            Text text73 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text73.Text = " ";

            run146.Append(runProperties146);
            run146.Append(text73);

            Run run147 = new Run() { RsidRunAddition = "00243370" };

            RunProperties runProperties147 = new RunProperties();
            RunFonts runFonts188 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold158 = new Bold();
            BoldComplexScript boldComplexScript84 = new BoldComplexScript();
            FontSize fontSize142 = new FontSize() { Val = "28" };

            runProperties147.Append(runFonts188);
            runProperties147.Append(bold158);
            runProperties147.Append(boldComplexScript84);
            runProperties147.Append(fontSize142);
            Text text74 = new Text();
            text74.Text = "regulator";

            run147.Append(runProperties147);
            run147.Append(text74);

            Run run148 = new Run() { RsidRunAddition = "00242440" };

            RunProperties runProperties148 = new RunProperties();
            RunFonts runFonts189 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold159 = new Bold();
            BoldComplexScript boldComplexScript85 = new BoldComplexScript();
            FontSize fontSize143 = new FontSize() { Val = "28" };

            runProperties148.Append(runFonts189);
            runProperties148.Append(bold159);
            runProperties148.Append(boldComplexScript85);
            runProperties148.Append(fontSize143);
            Text text75 = new Text();
            text75.Text = "y";

            run148.Append(runProperties148);
            run148.Append(text75);

            Run run149 = new Run() { RsidRunAddition = "00243370" };

            RunProperties runProperties149 = new RunProperties();
            RunFonts runFonts190 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold160 = new Bold();
            BoldComplexScript boldComplexScript86 = new BoldComplexScript();
            FontSize fontSize144 = new FontSize() { Val = "28" };

            runProperties149.Append(runFonts190);
            runProperties149.Append(bold160);
            runProperties149.Append(boldComplexScript86);
            runProperties149.Append(fontSize144);
            Text text76 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text76.Text = " ";

            run149.Append(runProperties149);
            run149.Append(text76);

            Run run150 = new Run() { RsidRunProperties = "00324D5C", RsidRunAddition = "00324D5C" };

            RunProperties runProperties150 = new RunProperties();
            RunFonts runFonts191 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold161 = new Bold();
            BoldComplexScript boldComplexScript87 = new BoldComplexScript();
            FontSize fontSize145 = new FontSize() { Val = "28" };

            runProperties150.Append(runFonts191);
            runProperties150.Append(bold161);
            runProperties150.Append(boldComplexScript87);
            runProperties150.Append(fontSize145);
            Text text77 = new Text();
            text77.Text = "enforcement action";

            run150.Append(runProperties150);
            run150.Append(text77);

            Run run151 = new Run() { RsidRunAddition = "00B52FB7" };

            RunProperties runProperties151 = new RunProperties();
            RunFonts runFonts192 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold162 = new Bold();
            BoldComplexScript boldComplexScript88 = new BoldComplexScript();
            FontSize fontSize146 = new FontSize() { Val = "28" };

            runProperties151.Append(runFonts192);
            runProperties151.Append(bold162);
            runProperties151.Append(boldComplexScript88);
            runProperties151.Append(fontSize146);
            Text text78 = new Text();
            text78.Text = ",";

            run151.Append(runProperties151);
            run151.Append(text78);

            Run run152 = new Run() { RsidRunAddition = "00B52FB7" };

            RunProperties runProperties152 = new RunProperties();
            RunFonts runFonts193 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold163 = new Bold();
            BoldComplexScript boldComplexScript89 = new BoldComplexScript();
            FontSize fontSize147 = new FontSize() { Val = "28" };

            runProperties152.Append(runFonts193);
            runProperties152.Append(bold163);
            runProperties152.Append(boldComplexScript89);
            runProperties152.Append(fontSize147);
            Text text79 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text79.Text = " major incident,";

            run152.Append(runProperties152);
            run152.Append(text79);

            Run run153 = new Run() { RsidRunAddition = "005C19FA" };

            RunProperties runProperties153 = new RunProperties();
            RunFonts runFonts194 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold164 = new Bold();
            BoldComplexScript boldComplexScript90 = new BoldComplexScript();
            FontSize fontSize148 = new FontSize() { Val = "28" };

            runProperties153.Append(runFonts194);
            runProperties153.Append(bold164);
            runProperties153.Append(boldComplexScript90);
            runProperties153.Append(fontSize148);
            Text text80 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text80.Text = " the";

            run153.Append(runProperties153);
            run153.Append(text80);

            Run run154 = new Run() { RsidRunAddition = "00B52FB7" };

            RunProperties runProperties154 = new RunProperties();
            RunFonts runFonts195 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold165 = new Bold();
            BoldComplexScript boldComplexScript91 = new BoldComplexScript();
            FontSize fontSize149 = new FontSize() { Val = "28" };

            runProperties154.Append(runFonts195);
            runProperties154.Append(bold165);
            runProperties154.Append(boldComplexScript91);
            runProperties154.Append(fontSize149);
            Text text81 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text81.Text = " ";

            run154.Append(runProperties154);
            run154.Append(text81);

            Run run155 = new Run() { RsidRunAddition = "00716ABD" };

            RunProperties runProperties155 = new RunProperties();
            RunFonts runFonts196 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold166 = new Bold();
            BoldComplexScript boldComplexScript92 = new BoldComplexScript();
            FontSize fontSize150 = new FontSize() { Val = "28" };

            runProperties155.Append(runFonts196);
            runProperties155.Append(bold166);
            runProperties155.Append(boldComplexScript92);
            runProperties155.Append(fontSize150);
            Text text82 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text82.Text = "compliance ";

            run155.Append(runProperties155);
            run155.Append(text82);

            Run run156 = new Run() { RsidRunAddition = "005C19FA" };

            RunProperties runProperties156 = new RunProperties();
            RunFonts runFonts197 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold167 = new Bold();
            BoldComplexScript boldComplexScript93 = new BoldComplexScript();
            FontSize fontSize151 = new FontSize() { Val = "28" };

            runProperties156.Append(runFonts197);
            runProperties156.Append(bold167);
            runProperties156.Append(boldComplexScript93);
            runProperties156.Append(fontSize151);
            Text text83 = new Text();
            text83.Text = "of";

            run156.Append(runProperties156);
            run156.Append(text83);

            Run run157 = new Run() { RsidRunAddition = "006A5363" };

            RunProperties runProperties157 = new RunProperties();
            RunFonts runFonts198 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold168 = new Bold();
            BoldComplexScript boldComplexScript94 = new BoldComplexScript();
            FontSize fontSize152 = new FontSize() { Val = "28" };

            runProperties157.Append(runFonts198);
            runProperties157.Append(bold168);
            runProperties157.Append(boldComplexScript94);
            runProperties157.Append(fontSize152);
            Text text84 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text84.Text = " ";

            run157.Append(runProperties157);
            run157.Append(text84);

            Run run158 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties158 = new RunProperties();
            RunFonts runFonts199 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold169 = new Bold();
            BoldComplexScript boldComplexScript95 = new BoldComplexScript();
            FontSize fontSize153 = new FontSize() { Val = "28" };

            runProperties158.Append(runFonts199);
            runProperties158.Append(bold169);
            runProperties158.Append(boldComplexScript95);
            runProperties158.Append(fontSize153);
            Text text85 = new Text();
            text85.Text = "competent";

            run158.Append(runProperties158);
            run158.Append(text85);

            Run run159 = new Run() { RsidRunAddition = "006A5363" };

            RunProperties runProperties159 = new RunProperties();
            RunFonts runFonts200 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold170 = new Bold();
            BoldComplexScript boldComplexScript96 = new BoldComplexScript();
            FontSize fontSize154 = new FontSize() { Val = "28" };

            runProperties159.Append(runFonts200);
            runProperties159.Append(bold170);
            runProperties159.Append(boldComplexScript96);
            runProperties159.Append(fontSize154);
            Text text86 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text86.Text = " authority";

            run159.Append(runProperties159);
            run159.Append(text86);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run160 = new Run() { RsidRunAddition = "006A5363" };

            RunProperties runProperties160 = new RunProperties();
            RunFonts runFonts201 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold171 = new Bold();
            BoldComplexScript boldComplexScript97 = new BoldComplexScript();
            FontSize fontSize155 = new FontSize() { Val = "28" };

            runProperties160.Append(runFonts201);
            runProperties160.Append(bold171);
            runProperties160.Append(boldComplexScript97);
            runProperties160.Append(fontSize155);
            Text text87 = new Text();
            text87.Text = "’";

            run160.Append(runProperties160);
            run160.Append(text87);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run161 = new Run() { RsidRunAddition = "006A5363" };

            RunProperties runProperties161 = new RunProperties();
            RunFonts runFonts202 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold172 = new Bold();
            BoldComplexScript boldComplexScript98 = new BoldComplexScript();
            FontSize fontSize156 = new FontSize() { Val = "28" };

            runProperties161.Append(runFonts202);
            runProperties161.Append(bold172);
            runProperties161.Append(boldComplexScript98);
            runProperties161.Append(fontSize156);
            Text text88 = new Text();
            text88.Text = "s directives";

            run161.Append(runProperties161);
            run161.Append(text88);

            Run run162 = new Run() { RsidRunAddition = "00242440" };

            RunProperties runProperties162 = new RunProperties();
            RunFonts runFonts203 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold173 = new Bold();
            BoldComplexScript boldComplexScript99 = new BoldComplexScript();
            FontSize fontSize157 = new FontSize() { Val = "28" };

            runProperties162.Append(runFonts203);
            runProperties162.Append(bold173);
            runProperties162.Append(boldComplexScript99);
            runProperties162.Append(fontSize157);
            Text text89 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text89.Text = " rega";

            run162.Append(runProperties162);
            run162.Append(text89);

            Run run163 = new Run() { RsidRunAddition = "00D66E75" };

            RunProperties runProperties163 = new RunProperties();
            RunFonts runFonts204 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold174 = new Bold();
            BoldComplexScript boldComplexScript100 = new BoldComplexScript();
            FontSize fontSize158 = new FontSize() { Val = "28" };

            runProperties163.Append(runFonts204);
            runProperties163.Append(bold174);
            runProperties163.Append(boldComplexScript100);
            runProperties163.Append(fontSize158);
            Text text90 = new Text();
            text90.Text = "r";

            run163.Append(runProperties163);
            run163.Append(text90);

            Run run164 = new Run() { RsidRunAddition = "00242440" };

            RunProperties runProperties164 = new RunProperties();
            RunFonts runFonts205 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold175 = new Bold();
            BoldComplexScript boldComplexScript101 = new BoldComplexScript();
            FontSize fontSize159 = new FontSize() { Val = "28" };

            runProperties164.Append(runFonts205);
            runProperties164.Append(bold175);
            runProperties164.Append(boldComplexScript101);
            runProperties164.Append(fontSize159);
            Text text91 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text91.Text = "ding ";

            run164.Append(runProperties164);
            run164.Append(text91);

            Run run165 = new Run() { RsidRunAddition = "00C413B0" };

            RunProperties runProperties165 = new RunProperties();
            RunFonts runFonts206 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold176 = new Bold();
            BoldComplexScript boldComplexScript102 = new BoldComplexScript();
            FontSize fontSize160 = new FontSize() { Val = "28" };

            runProperties165.Append(runFonts206);
            runProperties165.Append(bold176);
            runProperties165.Append(boldComplexScript102);
            runProperties165.Append(fontSize160);
            Text text92 = new Text();
            text92.Text = "bank";

            run165.Append(runProperties165);
            run165.Append(text92);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run166 = new Run() { RsidRunAddition = "00C413B0" };

            RunProperties runProperties166 = new RunProperties();
            RunFonts runFonts207 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold177 = new Bold();
            BoldComplexScript boldComplexScript103 = new BoldComplexScript();
            FontSize fontSize161 = new FontSize() { Val = "28" };

            runProperties166.Append(runFonts207);
            runProperties166.Append(bold177);
            runProperties166.Append(boldComplexScript103);
            runProperties166.Append(fontSize161);
            Text text93 = new Text();
            text93.Text = "’";

            run166.Append(runProperties166);
            run166.Append(text93);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run167 = new Run() { RsidRunAddition = "00C413B0" };

            RunProperties runProperties167 = new RunProperties();
            RunFonts runFonts208 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold178 = new Bold();
            BoldComplexScript boldComplexScript104 = new BoldComplexScript();
            FontSize fontSize162 = new FontSize() { Val = "28" };

            runProperties167.Append(runFonts208);
            runProperties167.Append(bold178);
            runProperties167.Append(boldComplexScript104);
            runProperties167.Append(fontSize162);
            Text text94 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text94.Text = "s ";

            run167.Append(runProperties167);
            run167.Append(text94);

            Run run168 = new Run() { RsidRunAddition = "00242440" };

            RunProperties runProperties168 = new RunProperties();
            RunFonts runFonts209 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold179 = new Bold();
            BoldComplexScript boldComplexScript105 = new BoldComplexScript();
            FontSize fontSize163 = new FontSize() { Val = "28" };

            runProperties168.Append(runFonts209);
            runProperties168.Append(bold179);
            runProperties168.Append(boldComplexScript105);
            runProperties168.Append(fontSize163);
            Text text95 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text95.Text = "business application or ";

            run168.Append(runProperties168);
            run168.Append(text95);

            Run run169 = new Run() { RsidRunAddition = "008D1E32" };

            RunProperties runProperties169 = new RunProperties();
            RunFonts runFonts210 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold180 = new Bold();
            BoldComplexScript boldComplexScript106 = new BoldComplexScript();
            FontSize fontSize164 = new FontSize() { Val = "28" };

            runProperties169.Append(runFonts210);
            runProperties169.Append(bold180);
            runProperties169.Append(boldComplexScript106);
            runProperties169.Append(fontSize164);
            Text text96 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text96.Text = "improvement measures, ";

            run169.Append(runProperties169);
            run169.Append(text96);

            Run run170 = new Run() { RsidRunAddition = "00F4377B" };

            RunProperties runProperties170 = new RunProperties();
            RunFonts runFonts211 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold181 = new Bold();
            BoldComplexScript boldComplexScript107 = new BoldComplexScript();
            FontSize fontSize165 = new FontSize() { Val = "28" };

            runProperties170.Append(runFonts211);
            runProperties170.Append(bold181);
            runProperties170.Append(boldComplexScript107);
            runProperties170.Append(fontSize165);
            Text text97 = new Text();
            text97.Text = "a";

            run170.Append(runProperties170);
            run170.Append(text97);

            Run run171 = new Run() { RsidRunAddition = "00F4377B" };

            RunProperties runProperties171 = new RunProperties();
            RunFonts runFonts212 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold182 = new Bold();
            BoldComplexScript boldComplexScript108 = new BoldComplexScript();
            FontSize fontSize166 = new FontSize() { Val = "28" };

            runProperties171.Append(runFonts212);
            runProperties171.Append(bold182);
            runProperties171.Append(boldComplexScript108);
            runProperties171.Append(fontSize166);
            Text text98 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text98.Text = "bnormal customer ";

            run171.Append(runProperties171);
            run171.Append(text98);

            Run run172 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties172 = new RunProperties();
            RunFonts runFonts213 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold183 = new Bold();
            BoldComplexScript boldComplexScript109 = new BoldComplexScript();
            FontSize fontSize167 = new FontSize() { Val = "28" };

            runProperties172.Append(runFonts213);
            runProperties172.Append(bold183);
            runProperties172.Append(boldComplexScript109);
            runProperties172.Append(fontSize167);
            Text text99 = new Text();
            text99.Text = "complaints";

            run172.Append(runProperties172);
            run172.Append(text99);

            Run run173 = new Run() { RsidRunAddition = "00F4377B" };

            RunProperties runProperties173 = new RunProperties();
            RunFonts runFonts214 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold184 = new Bold();
            BoldComplexScript boldComplexScript110 = new BoldComplexScript();
            FontSize fontSize168 = new FontSize() { Val = "28" };

            runProperties173.Append(runFonts214);
            runProperties173.Append(bold184);
            runProperties173.Append(boldComplexScript110);
            runProperties173.Append(fontSize168);
            Text text100 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text100.Text = " regarding structured products)";

            run173.Append(runProperties173);
            run173.Append(text100);

            Run run174 = new Run() { RsidRunProperties = "00B3199E", RsidRunAddition = "00F03C81" };

            RunProperties runProperties174 = new RunProperties();
            RunFonts runFonts215 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold185 = new Bold();
            VerticalTextAlignment verticalTextAlignment2 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            runProperties174.Append(runFonts215);
            runProperties174.Append(bold185);
            runProperties174.Append(verticalTextAlignment2);
            FootnoteReference footnoteReference3 = new FootnoteReference() { Id = 3 };

            run174.Append(runProperties174);
            run174.Append(footnoteReference3);

            Run run175 = new Run() { RsidRunProperties = "00D02A33", RsidRunAddition = "0014137D" };

            RunProperties runProperties175 = new RunProperties();
            RunFonts runFonts216 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold186 = new Bold();
            BoldComplexScript boldComplexScript111 = new BoldComplexScript();
            FontSize fontSize169 = new FontSize() { Val = "28" };

            runProperties175.Append(runFonts216);
            runProperties175.Append(bold186);
            runProperties175.Append(boldComplexScript111);
            runProperties175.Append(fontSize169);
            Text text101 = new Text();
            text101.Text = "：";

            run175.Append(runProperties175);
            run175.Append(text101);

            paragraph41.Append(paragraphProperties41);
            paragraph41.Append(run140);
            paragraph41.Append(run141);
            paragraph41.Append(run142);
            paragraph41.Append(run143);
            paragraph41.Append(run144);
            paragraph41.Append(run145);
            paragraph41.Append(run146);
            paragraph41.Append(run147);
            paragraph41.Append(run148);
            paragraph41.Append(run149);
            paragraph41.Append(run150);
            paragraph41.Append(run151);
            paragraph41.Append(run152);
            paragraph41.Append(run153);
            paragraph41.Append(run154);
            paragraph41.Append(run155);
            paragraph41.Append(run156);
            paragraph41.Append(run157);
            paragraph41.Append(run158);
            paragraph41.Append(run159);
            paragraph41.Append(proofError1);
            paragraph41.Append(run160);
            paragraph41.Append(proofError2);
            paragraph41.Append(run161);
            paragraph41.Append(run162);
            paragraph41.Append(run163);
            paragraph41.Append(run164);
            paragraph41.Append(run165);
            paragraph41.Append(proofError3);
            paragraph41.Append(run166);
            paragraph41.Append(proofError4);
            paragraph41.Append(run167);
            paragraph41.Append(run168);
            paragraph41.Append(run169);
            paragraph41.Append(run170);
            paragraph41.Append(run171);
            paragraph41.Append(run172);
            paragraph41.Append(run173);
            paragraph41.Append(run174);
            paragraph41.Append(run175);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00FE4552", RsidParagraphProperties = "004C41E7", RsidRunAdditionDefault = "00800493", ParagraphId = "6D8D64F6", TextId = "6A760BB6" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();

            NumberingProperties numberingProperties9 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference9 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId9 = new NumberingId() { Val = 15 };

            numberingProperties9.Append(numberingLevelReference9);
            numberingProperties9.Append(numberingId9);
            SnapToGrid snapToGrid38 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            RunFonts runFonts217 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold187 = new Bold();
            BoldComplexScript boldComplexScript112 = new BoldComplexScript();

            paragraphMarkRunProperties42.Append(runFonts217);
            paragraphMarkRunProperties42.Append(bold187);
            paragraphMarkRunProperties42.Append(boldComplexScript112);

            paragraphProperties42.Append(numberingProperties9);
            paragraphProperties42.Append(snapToGrid38);
            paragraphProperties42.Append(paragraphMarkRunProperties42);

            Run run176 = new Run();

            RunProperties runProperties176 = new RunProperties();
            RunFonts runFonts218 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold188 = new Bold();
            BoldComplexScript boldComplexScript113 = new BoldComplexScript();
            FontSize fontSize170 = new FontSize() { Val = "28" };

            runProperties176.Append(runFonts218);
            runProperties176.Append(bold188);
            runProperties176.Append(boldComplexScript113);
            runProperties176.Append(fontSize170);
            Text text102 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text102.Text = "Collaborative auditing with ";

            run176.Append(runProperties176);
            run176.Append(text102);

            Run run177 = new Run();

            RunProperties runProperties177 = new RunProperties();
            RunFonts runFonts219 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold189 = new Bold();
            BoldComplexScript boldComplexScript114 = new BoldComplexScript();
            FontSize fontSize171 = new FontSize() { Val = "28" };

            runProperties177.Append(runFonts219);
            runProperties177.Append(bold189);
            runProperties177.Append(boldComplexScript114);
            runProperties177.Append(fontSize171);
            Text text103 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text103.Text = "head office IA or other Location ";

            run177.Append(runProperties177);
            run177.Append(text103);

            Run run178 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties178 = new RunProperties();
            RunFonts runFonts220 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold190 = new Bold();
            BoldComplexScript boldComplexScript115 = new BoldComplexScript();
            FontSize fontSize172 = new FontSize() { Val = "28" };

            runProperties178.Append(runFonts220);
            runProperties178.Append(bold190);
            runProperties178.Append(boldComplexScript115);
            runProperties178.Append(fontSize172);
            Text text104 = new Text();
            text104.Text = "IA (";

            run178.Append(runProperties178);
            run178.Append(text104);

            Run run179 = new Run() { RsidRunAddition = "00013384" };

            RunProperties runProperties179 = new RunProperties();
            RunFonts runFonts221 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold191 = new Bold();
            BoldComplexScript boldComplexScript116 = new BoldComplexScript();
            FontSize fontSize173 = new FontSize() { Val = "28" };

            runProperties179.Append(runFonts221);
            runProperties179.Append(bold191);
            runProperties179.Append(boldComplexScript116);
            runProperties179.Append(fontSize173);
            Text text105 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text105.Text = "please ";

            run179.Append(runProperties179);
            run179.Append(text105);

            Run run180 = new Run() { RsidRunProperties = "004C41E7", RsidRunAddition = "004C41E7" };

            RunProperties runProperties180 = new RunProperties();
            RunFonts runFonts222 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold192 = new Bold();
            BoldComplexScript boldComplexScript117 = new BoldComplexScript();
            FontSize fontSize174 = new FontSize() { Val = "28" };

            runProperties180.Append(runFonts222);
            runProperties180.Append(bold192);
            runProperties180.Append(boldComplexScript117);
            runProperties180.Append(fontSize174);
            Text text106 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text106.Text = "elaborate ";

            run180.Append(runProperties180);
            run180.Append(text106);

            Run run181 = new Run() { RsidRunAddition = "00013384" };

            RunProperties runProperties181 = new RunProperties();
            RunFonts runFonts223 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold193 = new Bold();
            BoldComplexScript boldComplexScript118 = new BoldComplexScript();
            FontSize fontSize175 = new FontSize() { Val = "28" };

            runProperties181.Append(runFonts223);
            runProperties181.Append(bold193);
            runProperties181.Append(boldComplexScript118);
            runProperties181.Append(fontSize175);
            Text text107 = new Text();
            text107.Text = "the methods used to conduct this collaborative audit and the respective responsibilities assigned to each party";

            run181.Append(runProperties181);
            run181.Append(text107);

            Run run182 = new Run() { RsidRunProperties = "00787422", RsidRunAddition = "00CF43E8" };

            RunProperties runProperties182 = new RunProperties();
            RunFonts runFonts224 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold194 = new Bold();
            BoldComplexScript boldComplexScript119 = new BoldComplexScript();
            FontSize fontSize176 = new FontSize() { Val = "28" };

            runProperties182.Append(runFonts224);
            runProperties182.Append(bold194);
            runProperties182.Append(boldComplexScript119);
            runProperties182.Append(fontSize176);
            Text text108 = new Text();
            text108.Text = ")";

            run182.Append(runProperties182);
            run182.Append(text108);

            paragraph42.Append(paragraphProperties42);
            paragraph42.Append(run176);
            paragraph42.Append(run177);
            paragraph42.Append(run178);
            paragraph42.Append(run179);
            paragraph42.Append(run180);
            paragraph42.Append(run181);
            paragraph42.Append(run182);

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "003619F7", RsidParagraphProperties = "003619F7", RsidRunAdditionDefault = "003619F7", ParagraphId = "646722F5", TextId = "77777777" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();
            SnapToGrid snapToGrid39 = new SnapToGrid() { Val = false };
            Indentation indentation21 = new Indentation() { Start = "480" };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            RunFonts runFonts225 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold195 = new Bold();
            BoldComplexScript boldComplexScript120 = new BoldComplexScript();

            paragraphMarkRunProperties43.Append(runFonts225);
            paragraphMarkRunProperties43.Append(bold195);
            paragraphMarkRunProperties43.Append(boldComplexScript120);

            paragraphProperties43.Append(snapToGrid39);
            paragraphProperties43.Append(indentation21);
            paragraphProperties43.Append(paragraphMarkRunProperties43);

            paragraph43.Append(paragraphProperties43);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "003619F7", RsidParagraphAddition = "00882386", RsidParagraphProperties = "003619F7", RsidRunAdditionDefault = "003619F7", ParagraphId = "63B4F8ED", TextId = "77777777" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();

            NumberingProperties numberingProperties10 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference10 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId10 = new NumberingId() { Val = 15 };

            numberingProperties10.Append(numberingLevelReference10);
            numberingProperties10.Append(numberingId10);
            SnapToGrid snapToGrid40 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            RunFonts runFonts226 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold196 = new Bold();
            BoldComplexScript boldComplexScript121 = new BoldComplexScript();
            FontSize fontSize177 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties44.Append(runFonts226);
            paragraphMarkRunProperties44.Append(bold196);
            paragraphMarkRunProperties44.Append(boldComplexScript121);
            paragraphMarkRunProperties44.Append(fontSize177);

            paragraphProperties44.Append(numberingProperties10);
            paragraphProperties44.Append(snapToGrid40);
            paragraphProperties44.Append(paragraphMarkRunProperties44);

            Run run183 = new Run();

            RunProperties runProperties183 = new RunProperties();
            RunFonts runFonts227 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold197 = new Bold();
            BoldComplexScript boldComplexScript122 = new BoldComplexScript();
            FontSize fontSize178 = new FontSize() { Val = "28" };

            runProperties183.Append(runFonts227);
            runProperties183.Append(bold197);
            runProperties183.Append(boldComplexScript122);
            runProperties183.Append(fontSize178);
            Text text109 = new Text();
            text109.Text = "O";

            run183.Append(runProperties183);
            run183.Append(text109);

            Run run184 = new Run();

            RunProperties runProperties184 = new RunProperties();
            RunFonts runFonts228 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold198 = new Bold();
            BoldComplexScript boldComplexScript123 = new BoldComplexScript();
            FontSize fontSize179 = new FontSize() { Val = "28" };

            runProperties184.Append(runFonts228);
            runProperties184.Append(bold198);
            runProperties184.Append(boldComplexScript123);
            runProperties184.Append(fontSize179);
            Text text110 = new Text();
            text110.Text = "thers:";

            run184.Append(runProperties184);
            run184.Append(text110);

            paragraph44.Append(paragraphProperties44);
            paragraph44.Append(run183);
            paragraph44.Append(run184);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "00213B89", ParagraphId = "30BE99D4", TextId = "77777777" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            RunFonts runFonts229 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold199 = new Bold();
            BoldComplexScript boldComplexScript124 = new BoldComplexScript();
            FontSize fontSize180 = new FontSize() { Val = "40" };
            Underline underline4 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties45.Append(runFonts229);
            paragraphMarkRunProperties45.Append(bold199);
            paragraphMarkRunProperties45.Append(boldComplexScript124);
            paragraphMarkRunProperties45.Append(fontSize180);
            paragraphMarkRunProperties45.Append(underline4);

            paragraphProperties45.Append(paragraphStyleId3);
            paragraphProperties45.Append(spacingBetweenLines5);
            paragraphProperties45.Append(paragraphMarkRunProperties45);

            paragraph45.Append(paragraphProperties45);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "00213B89", ParagraphId = "1B202BEA", TextId = "77777777" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            RunFonts runFonts230 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold200 = new Bold();
            BoldComplexScript boldComplexScript125 = new BoldComplexScript();
            FontSize fontSize181 = new FontSize() { Val = "40" };
            Underline underline5 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties46.Append(runFonts230);
            paragraphMarkRunProperties46.Append(bold200);
            paragraphMarkRunProperties46.Append(boldComplexScript125);
            paragraphMarkRunProperties46.Append(fontSize181);
            paragraphMarkRunProperties46.Append(underline5);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidRPr = "0096027D", RsidR = "00213B89", RsidSect = "00213B89" };
            HeaderReference headerReference1 = new HeaderReference() { Type = HeaderFooterValues.Default, Id = "rId11" };
            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId12" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin1 = new PageMargin() { Top = 1701, Right = (UInt32Value)1077U, Bottom = 1440, Left = (UInt32Value)1077U, Header = (UInt32Value)680U, Footer = (UInt32Value)567U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "425" };
            DocGrid docGrid1 = new DocGrid() { Type = DocGridValues.Lines, LinePitch = 360 };

            sectionProperties1.Append(headerReference1);
            sectionProperties1.Append(footerReference1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            paragraphProperties46.Append(paragraphStyleId4);
            paragraphProperties46.Append(spacingBetweenLines6);
            paragraphProperties46.Append(paragraphMarkRunProperties46);
            paragraphProperties46.Append(sectionProperties1);

            paragraph46.Append(paragraphProperties46);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00200048", RsidRunAdditionDefault = "009023C8", ParagraphId = "3C4177F1", TextId = "719063A7" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification29 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            RunFonts runFonts231 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold201 = new Bold();
            BoldComplexScript boldComplexScript126 = new BoldComplexScript();
            FontSize fontSize182 = new FontSize() { Val = "40" };
            Underline underline6 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties47.Append(runFonts231);
            paragraphMarkRunProperties47.Append(bold201);
            paragraphMarkRunProperties47.Append(boldComplexScript126);
            paragraphMarkRunProperties47.Append(fontSize182);
            paragraphMarkRunProperties47.Append(underline6);

            paragraphProperties47.Append(paragraphStyleId5);
            paragraphProperties47.Append(spacingBetweenLines7);
            paragraphProperties47.Append(justification29);
            paragraphProperties47.Append(paragraphMarkRunProperties47);

            Run run185 = new Run() { RsidRunProperties = "009023C8" };

            RunProperties runProperties185 = new RunProperties();
            RunFonts runFonts232 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold202 = new Bold();
            BoldComplexScript boldComplexScript127 = new BoldComplexScript();
            FontSize fontSize183 = new FontSize() { Val = "40" };
            Underline underline7 = new Underline() { Val = UnderlineValues.Single };

            runProperties185.Append(runFonts232);
            runProperties185.Append(bold202);
            runProperties185.Append(boldComplexScript127);
            runProperties185.Append(fontSize183);
            runProperties185.Append(underline7);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text111 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text111.Text = "Engagement ";

            run185.Append(runProperties185);
            run185.Append(lastRenderedPageBreak1);
            run185.Append(text111);

            Run run186 = new Run() { RsidRunProperties = "009023C8", RsidRunAddition = "00153D64" };

            RunProperties runProperties186 = new RunProperties();
            RunFonts runFonts233 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold203 = new Bold();
            BoldComplexScript boldComplexScript128 = new BoldComplexScript();
            FontSize fontSize184 = new FontSize() { Val = "40" };
            Underline underline8 = new Underline() { Val = UnderlineValues.Single };

            runProperties186.Append(runFonts233);
            runProperties186.Append(bold203);
            runProperties186.Append(boldComplexScript128);
            runProperties186.Append(fontSize184);
            runProperties186.Append(underline8);
            Text text112 = new Text();
            text112.Text = "Plan";

            run186.Append(runProperties186);
            run186.Append(text112);

            Run run187 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties187 = new RunProperties();
            RunFonts runFonts234 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold204 = new Bold();
            BoldComplexScript boldComplexScript129 = new BoldComplexScript();
            FontSize fontSize185 = new FontSize() { Val = "40" };
            Underline underline9 = new Underline() { Val = UnderlineValues.Single };

            runProperties187.Append(runFonts234);
            runProperties187.Append(bold204);
            runProperties187.Append(boldComplexScript129);
            runProperties187.Append(fontSize185);
            runProperties187.Append(underline9);
            Text text113 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text113.Text = " (";

            run187.Append(runProperties187);
            run187.Append(text113);

            Run run188 = new Run();

            RunProperties runProperties188 = new RunProperties();
            RunFonts runFonts235 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold205 = new Bold();
            BoldComplexScript boldComplexScript130 = new BoldComplexScript();
            FontSize fontSize186 = new FontSize() { Val = "40" };
            Underline underline10 = new Underline() { Val = UnderlineValues.Single };

            runProperties188.Append(runFonts235);
            runProperties188.Append(bold205);
            runProperties188.Append(boldComplexScript130);
            runProperties188.Append(fontSize186);
            runProperties188.Append(underline10);
            Text text114 = new Text();
            text114.Text = "Thematic audit)";

            run188.Append(runProperties188);
            run188.Append(text114);

            paragraph47.Append(paragraphProperties47);
            paragraph47.Append(run185);
            paragraph47.Append(run186);
            paragraph47.Append(run187);
            paragraph47.Append(run188);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "003B1BF1", RsidParagraphProperties = "00200048", RsidRunAdditionDefault = "003B1BF1", ParagraphId = "36A31A6A", TextId = "77777777" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification30 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            RunFonts runFonts236 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold206 = new Bold();
            BoldComplexScript boldComplexScript131 = new BoldComplexScript();
            FontSize fontSize187 = new FontSize() { Val = "40" };
            Underline underline11 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties48.Append(runFonts236);
            paragraphMarkRunProperties48.Append(bold206);
            paragraphMarkRunProperties48.Append(boldComplexScript131);
            paragraphMarkRunProperties48.Append(fontSize187);
            paragraphMarkRunProperties48.Append(underline11);

            paragraphProperties48.Append(paragraphStyleId6);
            paragraphProperties48.Append(spacingBetweenLines8);
            paragraphProperties48.Append(justification30);
            paragraphProperties48.Append(paragraphMarkRunProperties48);

            paragraph48.Append(paragraphProperties48);

            Table table3 = new Table();

            TableProperties tableProperties3 = new TableProperties();
            TableWidth tableWidth3 = new TableWidth() { Width = "10400", Type = TableWidthUnitValues.Dxa };
            TableIndentation tableIndentation3 = new TableIndentation() { Width = 28, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault3 = new TableCellMarginDefault();
            TableCellLeftMargin tableCellLeftMargin3 = new TableCellLeftMargin() { Width = 28, Type = TableWidthValues.Dxa };
            TableCellRightMargin tableCellRightMargin3 = new TableCellRightMargin() { Width = 28, Type = TableWidthValues.Dxa };

            tableCellMarginDefault3.Append(tableCellLeftMargin3);
            tableCellMarginDefault3.Append(tableCellRightMargin3);
            TableLook tableLook3 = new TableLook() { Val = "0000" };

            tableProperties3.Append(tableWidth3);
            tableProperties3.Append(tableIndentation3);
            tableProperties3.Append(tableCellMarginDefault3);
            tableProperties3.Append(tableLook3);

            TableGrid tableGrid3 = new TableGrid();
            GridColumn gridColumn6 = new GridColumn() { Width = "2410" };
            GridColumn gridColumn7 = new GridColumn() { Width = "7990" };

            tableGrid3.Append(gridColumn6);
            tableGrid3.Append(gridColumn7);

            TableRow tableRow12 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "7FC316AA", TextId = "77777777" };

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties27.Append(tableCellWidth27);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "2A392163", TextId = "77777777" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            SnapToGrid snapToGrid41 = new SnapToGrid() { Val = false };
            Indentation indentation22 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification31 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            RunFonts runFonts237 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold207 = new Bold();
            FontSize fontSize188 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties49.Append(runFonts237);
            paragraphMarkRunProperties49.Append(bold207);
            paragraphMarkRunProperties49.Append(fontSize188);

            paragraphProperties49.Append(snapToGrid41);
            paragraphProperties49.Append(indentation22);
            paragraphProperties49.Append(justification31);
            paragraphProperties49.Append(paragraphMarkRunProperties49);

            Run run189 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties189 = new RunProperties();
            RunFonts runFonts238 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold208 = new Bold();
            FontSize fontSize189 = new FontSize() { Val = "28" };
            Languages languages21 = new Languages() { EastAsia = "zh-HK" };

            runProperties189.Append(runFonts238);
            runProperties189.Append(bold208);
            runProperties189.Append(fontSize189);
            runProperties189.Append(languages21);
            Text text115 = new Text();
            text115.Text = "Audit Project";

            run189.Append(runProperties189);
            run189.Append(text115);

            Run run190 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties190 = new RunProperties();
            RunFonts runFonts239 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold209 = new Bold();
            FontSize fontSize190 = new FontSize() { Val = "28" };

            runProperties190.Append(runFonts239);
            runProperties190.Append(bold209);
            runProperties190.Append(fontSize190);
            Text text116 = new Text();
            text116.Text = ":";

            run190.Append(runProperties190);
            run190.Append(text116);

            paragraph49.Append(paragraphProperties49);
            paragraph49.Append(run189);
            paragraph49.Append(run190);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph49);

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties28.Append(tableCellWidth28);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "6CC29460", TextId = "1325734E" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            SnapToGrid snapToGrid42 = new SnapToGrid() { Val = false };
            Indentation indentation23 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification32 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            RunFonts runFonts240 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold210 = new Bold();
            BoldComplexScript boldComplexScript132 = new BoldComplexScript();
            Color color99 = new Color() { Val = "0000FF" };
            FontSize fontSize191 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "28" };

            paragraphMarkRunProperties50.Append(runFonts240);
            paragraphMarkRunProperties50.Append(bold210);
            paragraphMarkRunProperties50.Append(boldComplexScript132);
            paragraphMarkRunProperties50.Append(color99);
            paragraphMarkRunProperties50.Append(fontSize191);
            paragraphMarkRunProperties50.Append(fontSizeComplexScript41);

            paragraphProperties50.Append(snapToGrid42);
            paragraphProperties50.Append(indentation23);
            paragraphProperties50.Append(justification32);
            paragraphProperties50.Append(paragraphMarkRunProperties50);

            DeletedRun deletedRun42 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "57" };

            Run run191 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties191 = new RunProperties();
            RunFonts runFonts241 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold211 = new Bold();
            BoldComplexScript boldComplexScript133 = new BoldComplexScript();
            Color color100 = new Color() { Val = "0000FF" };
            FontSize fontSize192 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "28" };

            runProperties191.Append(runFonts241);
            runProperties191.Append(bold211);
            runProperties191.Append(boldComplexScript133);
            runProperties191.Append(color100);
            runProperties191.Append(fontSize192);
            runProperties191.Append(fontSizeComplexScript42);
            DeletedText deletedText72 = new DeletedText();
            deletedText72.Text = "[";

            run191.Append(runProperties191);
            run191.Append(deletedText72);

            Run run192 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties192 = new RunProperties();
            RunFonts runFonts242 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold212 = new Bold();
            BoldComplexScript boldComplexScript134 = new BoldComplexScript();
            Color color101 = new Color() { Val = "0000FF" };
            FontSize fontSize193 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "28" };

            runProperties192.Append(runFonts242);
            runProperties192.Append(bold212);
            runProperties192.Append(boldComplexScript134);
            runProperties192.Append(color101);
            runProperties192.Append(fontSize193);
            runProperties192.Append(fontSizeComplexScript43);
            DeletedText deletedText73 = new DeletedText();
            deletedText73.Text = "查程";

            run192.Append(runProperties192);
            run192.Append(deletedText73);

            Run run193 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties193 = new RunProperties();
            RunFonts runFonts243 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold213 = new Bold();
            BoldComplexScript boldComplexScript135 = new BoldComplexScript();
            Color color102 = new Color() { Val = "0000FF" };
            FontSize fontSize194 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "28" };

            runProperties193.Append(runFonts243);
            runProperties193.Append(bold213);
            runProperties193.Append(boldComplexScript135);
            runProperties193.Append(color102);
            runProperties193.Append(fontSize194);
            runProperties193.Append(fontSizeComplexScript44);
            DeletedText deletedText74 = new DeletedText();
            deletedText74.Text = "].[";

            run193.Append(runProperties193);
            run193.Append(deletedText74);

            deletedRun42.Append(run191);
            deletedRun42.Append(run192);
            deletedRun42.Append(run193);

            Run run194 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties194 = new RunProperties();
            RunFonts runFonts244 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold214 = new Bold();
            BoldComplexScript boldComplexScript136 = new BoldComplexScript();
            Color color103 = new Color() { Val = "0000FF" };
            FontSize fontSize195 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "28" };

            runProperties194.Append(runFonts244);
            runProperties194.Append(bold214);
            runProperties194.Append(boldComplexScript136);
            runProperties194.Append(color103);
            runProperties194.Append(fontSize195);
            runProperties194.Append(fontSizeComplexScript45);
            Text text117 = new Text();
            text117.Text = dt.Rows[0]["planname"].ToString();

            run194.Append(runProperties194);
            run194.Append(text117);

            DeletedRun deletedRun43 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "58" };

            Run run195 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "0096615A" };

            RunProperties runProperties195 = new RunProperties();
            RunFonts runFonts245 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold215 = new Bold();
            BoldComplexScript boldComplexScript137 = new BoldComplexScript();
            Color color104 = new Color() { Val = "0000FF" };
            FontSize fontSize196 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "28" };

            runProperties195.Append(runFonts245);
            runProperties195.Append(bold215);
            runProperties195.Append(boldComplexScript137);
            runProperties195.Append(color104);
            runProperties195.Append(fontSize196);
            runProperties195.Append(fontSizeComplexScript46);
            DeletedText deletedText75 = new DeletedText();
            deletedText75.Text = "_ENG";

            run195.Append(runProperties195);
            run195.Append(deletedText75);

            deletedRun43.Append(run195);

            DeletedRun deletedRun44 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "59" };

            Run run196 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties196 = new RunProperties();
            RunFonts runFonts246 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold216 = new Bold();
            BoldComplexScript boldComplexScript138 = new BoldComplexScript();
            Color color105 = new Color() { Val = "0000FF" };
            FontSize fontSize197 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "28" };

            runProperties196.Append(runFonts246);
            runProperties196.Append(bold216);
            runProperties196.Append(boldComplexScript138);
            runProperties196.Append(color105);
            runProperties196.Append(fontSize197);
            runProperties196.Append(fontSizeComplexScript47);
            DeletedText deletedText76 = new DeletedText();
            deletedText76.Text = "]";

            run196.Append(runProperties196);
            run196.Append(deletedText76);

            deletedRun44.Append(run196);

            paragraph50.Append(paragraphProperties50);
            paragraph50.Append(deletedRun42);
            paragraph50.Append(run194);
            paragraph50.Append(deletedRun43);
            paragraph50.Append(deletedRun44);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph50);

            tableRow12.Append(tableCell27);
            tableRow12.Append(tableCell28);

            TableRow tableRow13 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "1EBAE288", TextId = "77777777" };

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties29.Append(tableCellWidth29);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "7A497842", TextId = "77777777" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            SnapToGrid snapToGrid43 = new SnapToGrid() { Val = false };
            Indentation indentation24 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification33 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            RunFonts runFonts247 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold217 = new Bold();
            FontSize fontSize198 = new FontSize() { Val = "28" };
            Languages languages22 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties51.Append(runFonts247);
            paragraphMarkRunProperties51.Append(bold217);
            paragraphMarkRunProperties51.Append(fontSize198);
            paragraphMarkRunProperties51.Append(languages22);

            paragraphProperties51.Append(snapToGrid43);
            paragraphProperties51.Append(indentation24);
            paragraphProperties51.Append(justification33);
            paragraphProperties51.Append(paragraphMarkRunProperties51);

            Run run197 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties197 = new RunProperties();
            RunFonts runFonts248 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold218 = new Bold();
            FontSize fontSize199 = new FontSize() { Val = "28" };

            runProperties197.Append(runFonts248);
            runProperties197.Append(bold218);
            runProperties197.Append(fontSize199);
            Text text118 = new Text();
            text118.Text = "Auditee:";

            run197.Append(runProperties197);
            run197.Append(text118);

            paragraph51.Append(paragraphProperties51);
            paragraph51.Append(run197);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph51);

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties30.Append(tableCellWidth30);

            Paragraph paragraph52 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "6F976998", TextId = "3CF2A60E" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            SnapToGrid snapToGrid44 = new SnapToGrid() { Val = false };
            Indentation indentation25 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification34 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            RunFonts runFonts249 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold219 = new Bold();
            Color color106 = new Color() { Val = "0000FF" };
            FontSize fontSize200 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties52.Append(runFonts249);
            paragraphMarkRunProperties52.Append(bold219);
            paragraphMarkRunProperties52.Append(color106);
            paragraphMarkRunProperties52.Append(fontSize200);

            paragraphProperties52.Append(snapToGrid44);
            paragraphProperties52.Append(indentation25);
            paragraphProperties52.Append(justification34);
            paragraphProperties52.Append(paragraphMarkRunProperties52);

            DeletedRun deletedRun45 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "60" };

            Run run198 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties198 = new RunProperties();
            RunFonts runFonts250 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold220 = new Bold();
            Color color107 = new Color() { Val = "0000FF" };
            FontSize fontSize201 = new FontSize() { Val = "28" };

            runProperties198.Append(runFonts250);
            runProperties198.Append(bold220);
            runProperties198.Append(color107);
            runProperties198.Append(fontSize201);
            DeletedText deletedText77 = new DeletedText();
            deletedText77.Text = "UNION([";

            run198.Append(runProperties198);
            run198.Append(deletedText77);

            Run run199 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties199 = new RunProperties();
            RunFonts runFonts251 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold221 = new Bold();
            Color color108 = new Color() { Val = "0000FF" };
            FontSize fontSize202 = new FontSize() { Val = "28" };

            runProperties199.Append(runFonts251);
            runProperties199.Append(bold221);
            runProperties199.Append(color108);
            runProperties199.Append(fontSize202);
            DeletedText deletedText78 = new DeletedText();
            deletedText78.Text = "查程";

            run199.Append(runProperties199);
            run199.Append(deletedText78);

            Run run200 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties200 = new RunProperties();
            RunFonts runFonts252 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold222 = new Bold();
            Color color109 = new Color() { Val = "0000FF" };
            FontSize fontSize203 = new FontSize() { Val = "28" };

            runProperties200.Append(runFonts252);
            runProperties200.Append(bold222);
            runProperties200.Append(color109);
            runProperties200.Append(fontSize203);
            DeletedText deletedText79 = new DeletedText();
            deletedText79.Text = "].[";

            run200.Append(runProperties200);
            run200.Append(deletedText79);

            deletedRun45.Append(run198);
            deletedRun45.Append(run199);
            deletedRun45.Append(run200);

            Run run201 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties201 = new RunProperties();
            RunFonts runFonts253 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold223 = new Bold();
            Color color110 = new Color() { Val = "0000FF" };
            FontSize fontSize204 = new FontSize() { Val = "28" };

            runProperties201.Append(runFonts253);
            runProperties201.Append(bold223);
            runProperties201.Append(color110);
            runProperties201.Append(fontSize204);
            Text text119 = new Text();
            text119.Text = dt.Rows[0]["auditplandept"].ToString();

            run201.Append(runProperties201);
            run201.Append(text119);

            DeletedRun deletedRun46 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "61" };

            Run run202 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "0096615A" };

            RunProperties runProperties202 = new RunProperties();
            RunFonts runFonts254 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold224 = new Bold();
            Color color111 = new Color() { Val = "0000FF" };
            FontSize fontSize205 = new FontSize() { Val = "28" };

            runProperties202.Append(runFonts254);
            runProperties202.Append(bold224);
            runProperties202.Append(color111);
            runProperties202.Append(fontSize205);
            DeletedText deletedText80 = new DeletedText();
            deletedText80.Text = "_Eng";

            run202.Append(runProperties202);
            run202.Append(deletedText80);

            deletedRun46.Append(run202);

            DeletedRun deletedRun47 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "62" };

            Run run203 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties203 = new RunProperties();
            RunFonts runFonts255 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold225 = new Bold();
            Color color112 = new Color() { Val = "0000FF" };
            FontSize fontSize206 = new FontSize() { Val = "28" };

            runProperties203.Append(runFonts255);
            runProperties203.Append(bold225);
            runProperties203.Append(color112);
            runProperties203.Append(fontSize206);
            DeletedText deletedText81 = new DeletedText();
            deletedText81.Text = "])";

            run203.Append(runProperties203);
            run203.Append(deletedText81);

            deletedRun47.Append(run203);

            paragraph52.Append(paragraphProperties52);
            paragraph52.Append(deletedRun45);
            paragraph52.Append(run201);
            paragraph52.Append(deletedRun46);
            paragraph52.Append(deletedRun47);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph52);

            tableRow13.Append(tableCell29);
            tableRow13.Append(tableCell30);

            TableRow tableRow14 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "207948D7", TextId = "77777777" };

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties31.Append(tableCellWidth31);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "02914E0E", TextId = "77777777" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            SnapToGrid snapToGrid45 = new SnapToGrid() { Val = false };
            Indentation indentation26 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification35 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            RunFonts runFonts256 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold226 = new Bold();
            FontSize fontSize207 = new FontSize() { Val = "28" };
            Languages languages23 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties53.Append(runFonts256);
            paragraphMarkRunProperties53.Append(bold226);
            paragraphMarkRunProperties53.Append(fontSize207);
            paragraphMarkRunProperties53.Append(languages23);

            paragraphProperties53.Append(snapToGrid45);
            paragraphProperties53.Append(indentation26);
            paragraphProperties53.Append(justification35);
            paragraphProperties53.Append(paragraphMarkRunProperties53);

            Run run204 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties204 = new RunProperties();
            RunFonts runFonts257 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold227 = new Bold();
            FontSize fontSize208 = new FontSize() { Val = "28" };
            Languages languages24 = new Languages() { EastAsia = "zh-HK" };

            runProperties204.Append(runFonts257);
            runProperties204.Append(bold227);
            runProperties204.Append(fontSize208);
            runProperties204.Append(languages24);
            Text text120 = new Text();
            text120.Text = "Audit Type";

            run204.Append(runProperties204);
            run204.Append(text120);

            Run run205 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties205 = new RunProperties();
            RunFonts runFonts258 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold228 = new Bold();
            FontSize fontSize209 = new FontSize() { Val = "28" };

            runProperties205.Append(runFonts258);
            runProperties205.Append(bold228);
            runProperties205.Append(fontSize209);
            Text text121 = new Text();
            text121.Text = ":";

            run205.Append(runProperties205);
            run205.Append(text121);

            paragraph53.Append(paragraphProperties53);
            paragraph53.Append(run204);
            paragraph53.Append(run205);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph53);

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties32.Append(tableCellWidth32);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "1B6C2FFA", TextId = "3BCE262B" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            SnapToGrid snapToGrid46 = new SnapToGrid() { Val = false };
            Indentation indentation27 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification36 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            RunFonts runFonts259 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold229 = new Bold();
            Color color113 = new Color() { Val = "0000FF" };
            FontSize fontSize210 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties54.Append(runFonts259);
            paragraphMarkRunProperties54.Append(bold229);
            paragraphMarkRunProperties54.Append(color113);
            paragraphMarkRunProperties54.Append(fontSize210);

            paragraphProperties54.Append(snapToGrid46);
            paragraphProperties54.Append(indentation27);
            paragraphProperties54.Append(justification36);
            paragraphProperties54.Append(paragraphMarkRunProperties54);

            DeletedRun deletedRun48 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "63" };

            Run run206 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties206 = new RunProperties();
            RunFonts runFonts260 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold230 = new Bold();
            BoldComplexScript boldComplexScript139 = new BoldComplexScript();
            Color color114 = new Color() { Val = "0000FF" };
            FontSize fontSize211 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "28" };

            runProperties206.Append(runFonts260);
            runProperties206.Append(bold230);
            runProperties206.Append(boldComplexScript139);
            runProperties206.Append(color114);
            runProperties206.Append(fontSize211);
            runProperties206.Append(fontSizeComplexScript48);
            DeletedText deletedText82 = new DeletedText();
            deletedText82.Text = "[";

            run206.Append(runProperties206);
            run206.Append(deletedText82);

            Run run207 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties207 = new RunProperties();
            RunFonts runFonts261 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold231 = new Bold();
            BoldComplexScript boldComplexScript140 = new BoldComplexScript();
            Color color115 = new Color() { Val = "0000FF" };
            FontSize fontSize212 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "28" };

            runProperties207.Append(runFonts261);
            runProperties207.Append(bold231);
            runProperties207.Append(boldComplexScript140);
            runProperties207.Append(color115);
            runProperties207.Append(fontSize212);
            runProperties207.Append(fontSizeComplexScript49);
            DeletedText deletedText83 = new DeletedText();
            deletedText83.Text = "查程";

            run207.Append(runProperties207);
            run207.Append(deletedText83);

            Run run208 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties208 = new RunProperties();
            RunFonts runFonts262 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold232 = new Bold();
            BoldComplexScript boldComplexScript141 = new BoldComplexScript();
            Color color116 = new Color() { Val = "0000FF" };
            FontSize fontSize213 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "28" };

            runProperties208.Append(runFonts262);
            runProperties208.Append(bold232);
            runProperties208.Append(boldComplexScript141);
            runProperties208.Append(color116);
            runProperties208.Append(fontSize213);
            runProperties208.Append(fontSizeComplexScript50);
            DeletedText deletedText84 = new DeletedText();
            deletedText84.Text = "].[";

            run208.Append(runProperties208);
            run208.Append(deletedText84);

            deletedRun48.Append(run206);
            deletedRun48.Append(run207);
            deletedRun48.Append(run208);

            Run run209 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties209 = new RunProperties();
            RunFonts runFonts263 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold233 = new Bold();
            BoldComplexScript boldComplexScript142 = new BoldComplexScript();
            Color color117 = new Color() { Val = "0000FF" };
            FontSize fontSize214 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "28" };

            runProperties209.Append(runFonts263);
            runProperties209.Append(bold233);
            runProperties209.Append(boldComplexScript142);
            runProperties209.Append(color117);
            runProperties209.Append(fontSize214);
            runProperties209.Append(fontSizeComplexScript51);
            Text text122 = new Text();
            text122.Text = dt.Rows[0]["plantype"].ToString();

            run209.Append(runProperties209);
            run209.Append(text122);

            DeletedRun deletedRun49 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "64" };

            Run run210 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "0096615A" };

            RunProperties runProperties210 = new RunProperties();
            RunFonts runFonts264 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold234 = new Bold();
            BoldComplexScript boldComplexScript143 = new BoldComplexScript();
            Color color118 = new Color() { Val = "0000FF" };
            FontSize fontSize215 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "28" };

            runProperties210.Append(runFonts264);
            runProperties210.Append(bold234);
            runProperties210.Append(boldComplexScript143);
            runProperties210.Append(color118);
            runProperties210.Append(fontSize215);
            runProperties210.Append(fontSizeComplexScript52);
            DeletedText deletedText85 = new DeletedText();
            deletedText85.Text = "_ENG";

            run210.Append(runProperties210);
            run210.Append(deletedText85);

            deletedRun49.Append(run210);

            DeletedRun deletedRun50 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "65" };

            Run run211 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties211 = new RunProperties();
            RunFonts runFonts265 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold235 = new Bold();
            BoldComplexScript boldComplexScript144 = new BoldComplexScript();
            Color color119 = new Color() { Val = "0000FF" };
            FontSize fontSize216 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "28" };

            runProperties211.Append(runFonts265);
            runProperties211.Append(bold235);
            runProperties211.Append(boldComplexScript144);
            runProperties211.Append(color119);
            runProperties211.Append(fontSize216);
            runProperties211.Append(fontSizeComplexScript53);
            DeletedText deletedText86 = new DeletedText();
            deletedText86.Text = "]";

            run211.Append(runProperties211);
            run211.Append(deletedText86);

            deletedRun50.Append(run211);

            paragraph54.Append(paragraphProperties54);
            paragraph54.Append(deletedRun48);
            paragraph54.Append(run209);
            paragraph54.Append(deletedRun49);
            paragraph54.Append(deletedRun50);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph54);

            tableRow14.Append(tableCell31);
            tableRow14.Append(tableCell32);

            TableRow tableRow15 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "006C09C0", TextId = "77777777" };

            TableCell tableCell33 = new TableCell();

            TableCellProperties tableCellProperties33 = new TableCellProperties();
            TableCellWidth tableCellWidth33 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties33.Append(tableCellWidth33);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "67EC215E", TextId = "77777777" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            SnapToGrid snapToGrid47 = new SnapToGrid() { Val = false };
            Indentation indentation28 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification37 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            RunFonts runFonts266 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold236 = new Bold();
            FontSize fontSize217 = new FontSize() { Val = "28" };
            Languages languages25 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties55.Append(runFonts266);
            paragraphMarkRunProperties55.Append(bold236);
            paragraphMarkRunProperties55.Append(fontSize217);
            paragraphMarkRunProperties55.Append(languages25);

            paragraphProperties55.Append(snapToGrid47);
            paragraphProperties55.Append(indentation28);
            paragraphProperties55.Append(justification37);
            paragraphProperties55.Append(paragraphMarkRunProperties55);

            Run run212 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties212 = new RunProperties();
            RunFonts runFonts267 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold237 = new Bold();
            FontSize fontSize218 = new FontSize() { Val = "28" };
            Languages languages26 = new Languages() { EastAsia = "zh-HK" };

            runProperties212.Append(runFonts267);
            runProperties212.Append(bold237);
            runProperties212.Append(fontSize218);
            runProperties212.Append(languages26);
            Text text123 = new Text();
            text123.Text = "Audit Period:";

            run212.Append(runProperties212);
            run212.Append(text123);

            paragraph55.Append(paragraphProperties55);
            paragraph55.Append(run212);

            tableCell33.Append(tableCellProperties33);
            tableCell33.Append(paragraph55);

            TableCell tableCell34 = new TableCell();

            TableCellProperties tableCellProperties34 = new TableCellProperties();
            TableCellWidth tableCellWidth34 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties34.Append(tableCellWidth34);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "68C8F501", TextId = "26A211C4" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            SnapToGrid snapToGrid48 = new SnapToGrid() { Val = false };
            Indentation indentation29 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification38 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            RunFonts runFonts268 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold238 = new Bold();
            Color color120 = new Color() { Val = "0000FF" };
            FontSize fontSize219 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties56.Append(runFonts268);
            paragraphMarkRunProperties56.Append(bold238);
            paragraphMarkRunProperties56.Append(color120);
            paragraphMarkRunProperties56.Append(fontSize219);

            paragraphProperties56.Append(snapToGrid48);
            paragraphProperties56.Append(indentation29);
            paragraphProperties56.Append(justification38);
            paragraphProperties56.Append(paragraphMarkRunProperties56);

            DeletedRun deletedRun51 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "66" };

            Run run213 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties213 = new RunProperties();
            RunFonts runFonts269 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold239 = new Bold();
            Color color121 = new Color() { Val = "0000FF" };
            FontSize fontSize220 = new FontSize() { Val = "28" };

            runProperties213.Append(runFonts269);
            runProperties213.Append(bold239);
            runProperties213.Append(color121);
            runProperties213.Append(fontSize220);
            DeletedText deletedText87 = new DeletedText();
            deletedText87.Text = "[";

            run213.Append(runProperties213);
            run213.Append(deletedText87);

            Run run214 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties214 = new RunProperties();
            RunFonts runFonts270 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold240 = new Bold();
            Color color122 = new Color() { Val = "0000FF" };
            FontSize fontSize221 = new FontSize() { Val = "28" };

            runProperties214.Append(runFonts270);
            runProperties214.Append(bold240);
            runProperties214.Append(color122);
            runProperties214.Append(fontSize221);
            DeletedText deletedText88 = new DeletedText();
            deletedText88.Text = "查程";

            run214.Append(runProperties214);
            run214.Append(deletedText88);

            Run run215 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties215 = new RunProperties();
            RunFonts runFonts271 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold241 = new Bold();
            Color color123 = new Color() { Val = "0000FF" };
            FontSize fontSize222 = new FontSize() { Val = "28" };

            runProperties215.Append(runFonts271);
            runProperties215.Append(bold241);
            runProperties215.Append(color123);
            runProperties215.Append(fontSize222);
            DeletedText deletedText89 = new DeletedText();
            deletedText89.Text = "].[";

            run215.Append(runProperties215);
            run215.Append(deletedText89);

            deletedRun51.Append(run213);
            deletedRun51.Append(run214);
            deletedRun51.Append(run215);

            Run run216 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties216 = new RunProperties();
            RunFonts runFonts272 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold242 = new Bold();
            Color color124 = new Color() { Val = "0000FF" };
            FontSize fontSize223 = new FontSize() { Val = "28" };

            runProperties216.Append(runFonts272);
            runProperties216.Append(bold242);
            runProperties216.Append(color124);
            runProperties216.Append(fontSize223);
            Text text124 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["startdate"].ToString(), out DateTime date1))
            {
                text124.Text = date1.ToString("yyyy-MM-dd");
            }
            else
            {
                text124.Text = "";
            }

            run216.Append(runProperties216);
            run216.Append(text124);

            DeletedRun deletedRun52 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "67" };

            Run run217 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties217 = new RunProperties();
            RunFonts runFonts273 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold243 = new Bold();
            Color color125 = new Color() { Val = "0000FF" };
            FontSize fontSize224 = new FontSize() { Val = "28" };

            runProperties217.Append(runFonts273);
            runProperties217.Append(bold243);
            runProperties217.Append(color125);
            runProperties217.Append(fontSize224);
            DeletedText deletedText90 = new DeletedText();
            deletedText90.Text = "]";

            run217.Append(runProperties217);
            run217.Append(deletedText90);

            deletedRun52.Append(run217);

            Run run218 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties218 = new RunProperties();
            RunFonts runFonts274 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold244 = new Bold();
            Color color126 = new Color() { Val = "0000FF" };
            FontSize fontSize225 = new FontSize() { Val = "28" };

            runProperties218.Append(runFonts274);
            runProperties218.Append(bold244);
            runProperties218.Append(color126);
            runProperties218.Append(fontSize225);
            Text text125 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text125.Text = " ~ ";

            run218.Append(runProperties218);
            run218.Append(text125);

            DeletedRun deletedRun53 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "68" };

            Run run219 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties219 = new RunProperties();
            RunFonts runFonts275 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold245 = new Bold();
            Color color127 = new Color() { Val = "0000FF" };
            FontSize fontSize226 = new FontSize() { Val = "28" };

            runProperties219.Append(runFonts275);
            runProperties219.Append(bold245);
            runProperties219.Append(color127);
            runProperties219.Append(fontSize226);
            DeletedText deletedText91 = new DeletedText();
            deletedText91.Text = "[";

            run219.Append(runProperties219);
            run219.Append(deletedText91);

            Run run220 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties220 = new RunProperties();
            RunFonts runFonts276 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold246 = new Bold();
            Color color128 = new Color() { Val = "0000FF" };
            FontSize fontSize227 = new FontSize() { Val = "28" };

            runProperties220.Append(runFonts276);
            runProperties220.Append(bold246);
            runProperties220.Append(color128);
            runProperties220.Append(fontSize227);
            DeletedText deletedText92 = new DeletedText();
            deletedText92.Text = "查程";

            run220.Append(runProperties220);
            run220.Append(deletedText92);

            Run run221 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties221 = new RunProperties();
            RunFonts runFonts277 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold247 = new Bold();
            Color color129 = new Color() { Val = "0000FF" };
            FontSize fontSize228 = new FontSize() { Val = "28" };

            runProperties221.Append(runFonts277);
            runProperties221.Append(bold247);
            runProperties221.Append(color129);
            runProperties221.Append(fontSize228);
            DeletedText deletedText93 = new DeletedText();
            deletedText93.Text = "].[";

            run221.Append(runProperties221);
            run221.Append(deletedText93);

            deletedRun53.Append(run219);
            deletedRun53.Append(run220);
            deletedRun53.Append(run221);

            Run run222 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties222 = new RunProperties();
            RunFonts runFonts278 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold248 = new Bold();
            Color color130 = new Color() { Val = "0000FF" };
            FontSize fontSize229 = new FontSize() { Val = "28" };

            runProperties222.Append(runFonts278);
            runProperties222.Append(bold248);
            runProperties222.Append(color130);
            runProperties222.Append(fontSize229);
            Text text126 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["enddate"].ToString(), out DateTime date4))
            {
                text126.Text = date4.ToString("yyyy-MM-dd");
            }
            else
            {
                text126.Text = "";
            }

            run222.Append(runProperties222);
            run222.Append(text126);

            DeletedRun deletedRun54 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "69" };

            Run run223 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties223 = new RunProperties();
            RunFonts runFonts279 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold249 = new Bold();
            Color color131 = new Color() { Val = "0000FF" };
            FontSize fontSize230 = new FontSize() { Val = "28" };

            runProperties223.Append(runFonts279);
            runProperties223.Append(bold249);
            runProperties223.Append(color131);
            runProperties223.Append(fontSize230);
            DeletedText deletedText94 = new DeletedText();
            deletedText94.Text = "]";

            run223.Append(runProperties223);
            run223.Append(deletedText94);

            deletedRun54.Append(run223);

            paragraph56.Append(paragraphProperties56);
            paragraph56.Append(deletedRun51);
            paragraph56.Append(run216);
            paragraph56.Append(deletedRun52);
            paragraph56.Append(run218);
            paragraph56.Append(deletedRun53);
            paragraph56.Append(run222);
            paragraph56.Append(deletedRun54);

            tableCell34.Append(tableCellProperties34);
            tableCell34.Append(paragraph56);

            tableRow15.Append(tableCell33);
            tableRow15.Append(tableCell34);

            TableRow tableRow16 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "1F11A433", TextId = "77777777" };

            TableCell tableCell35 = new TableCell();

            TableCellProperties tableCellProperties35 = new TableCellProperties();
            TableCellWidth tableCellWidth35 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties35.Append(tableCellWidth35);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "4BBEAAD2", TextId = "77777777" };

            ParagraphProperties paragraphProperties57 = new ParagraphProperties();
            SnapToGrid snapToGrid49 = new SnapToGrid() { Val = false };
            Indentation indentation30 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification39 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties57 = new ParagraphMarkRunProperties();
            RunFonts runFonts280 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold250 = new Bold();
            FontSize fontSize231 = new FontSize() { Val = "28" };
            Languages languages27 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties57.Append(runFonts280);
            paragraphMarkRunProperties57.Append(bold250);
            paragraphMarkRunProperties57.Append(fontSize231);
            paragraphMarkRunProperties57.Append(languages27);

            paragraphProperties57.Append(snapToGrid49);
            paragraphProperties57.Append(indentation30);
            paragraphProperties57.Append(justification39);
            paragraphProperties57.Append(paragraphMarkRunProperties57);

            Run run224 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties224 = new RunProperties();
            RunFonts runFonts281 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold251 = new Bold();
            FontSize fontSize232 = new FontSize() { Val = "28" };
            Languages languages28 = new Languages() { EastAsia = "zh-HK" };

            runProperties224.Append(runFonts281);
            runProperties224.Append(bold251);
            runProperties224.Append(fontSize232);
            runProperties224.Append(languages28);
            Text text127 = new Text();
            text127.Text = "Scope Period:";

            run224.Append(runProperties224);
            run224.Append(text127);

            paragraph57.Append(paragraphProperties57);
            paragraph57.Append(run224);

            tableCell35.Append(tableCellProperties35);
            tableCell35.Append(paragraph57);

            TableCell tableCell36 = new TableCell();

            TableCellProperties tableCellProperties36 = new TableCellProperties();
            TableCellWidth tableCellWidth36 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties36.Append(tableCellWidth36);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "6CAC0A3C", TextId = "77777777" };

            ParagraphProperties paragraphProperties58 = new ParagraphProperties();
            SnapToGrid snapToGrid50 = new SnapToGrid() { Val = false };
            Indentation indentation31 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification40 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties58 = new ParagraphMarkRunProperties();
            RunFonts runFonts282 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold252 = new Bold();
            Color color132 = new Color() { Val = "0000FF" };
            FontSize fontSize233 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties58.Append(runFonts282);
            paragraphMarkRunProperties58.Append(bold252);
            paragraphMarkRunProperties58.Append(color132);
            paragraphMarkRunProperties58.Append(fontSize233);

            paragraphProperties58.Append(snapToGrid50);
            paragraphProperties58.Append(indentation31);
            paragraphProperties58.Append(justification40);
            paragraphProperties58.Append(paragraphMarkRunProperties58);

            DeletedRun deletedRun55 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "70" };

            Run run225 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties225 = new RunProperties();
            RunFonts runFonts283 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold253 = new Bold();
            Color color133 = new Color() { Val = "0000FF" };
            FontSize fontSize234 = new FontSize() { Val = "28" };

            runProperties225.Append(runFonts283);
            runProperties225.Append(bold253);
            runProperties225.Append(color133);
            runProperties225.Append(fontSize234);
            DeletedText deletedText95 = new DeletedText();
            deletedText95.Text = "[";

            run225.Append(runProperties225);
            run225.Append(deletedText95);

            Run run226 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties226 = new RunProperties();
            RunFonts runFonts284 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold254 = new Bold();
            Color color134 = new Color() { Val = "0000FF" };
            FontSize fontSize235 = new FontSize() { Val = "28" };

            runProperties226.Append(runFonts284);
            runProperties226.Append(bold254);
            runProperties226.Append(color134);
            runProperties226.Append(fontSize235);
            DeletedText deletedText96 = new DeletedText();
            deletedText96.Text = "查程";

            run226.Append(runProperties226);
            run226.Append(deletedText96);

            Run run227 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties227 = new RunProperties();
            RunFonts runFonts285 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold255 = new Bold();
            Color color135 = new Color() { Val = "0000FF" };
            FontSize fontSize236 = new FontSize() { Val = "28" };

            runProperties227.Append(runFonts285);
            runProperties227.Append(bold255);
            runProperties227.Append(color135);
            runProperties227.Append(fontSize236);
            DeletedText deletedText97 = new DeletedText();
            deletedText97.Text = "].[";

            run227.Append(runProperties227);
            run227.Append(deletedText97);

            deletedRun55.Append(run225);
            deletedRun55.Append(run226);
            deletedRun55.Append(run227);

            Run run228 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties228 = new RunProperties();
            RunFonts runFonts286 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold256 = new Bold();
            Color color136 = new Color() { Val = "0000FF" };
            FontSize fontSize237 = new FontSize() { Val = "28" };

            runProperties228.Append(runFonts286);
            runProperties228.Append(bold256);
            runProperties228.Append(color136);
            runProperties228.Append(fontSize237);
            Text text128 = new Text();
            text128.Text = "查核範圍起日";

            run228.Append(runProperties228);
            run228.Append(text128);

            DeletedRun deletedRun56 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "71" };

            Run run229 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties229 = new RunProperties();
            RunFonts runFonts287 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold257 = new Bold();
            Color color137 = new Color() { Val = "0000FF" };
            FontSize fontSize238 = new FontSize() { Val = "28" };

            runProperties229.Append(runFonts287);
            runProperties229.Append(bold257);
            runProperties229.Append(color137);
            runProperties229.Append(fontSize238);
            DeletedText deletedText98 = new DeletedText();
            deletedText98.Text = "]";

            run229.Append(runProperties229);
            run229.Append(deletedText98);

            deletedRun56.Append(run229);

            Run run230 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties230 = new RunProperties();
            RunFonts runFonts288 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold258 = new Bold();
            Color color138 = new Color() { Val = "0000FF" };
            FontSize fontSize239 = new FontSize() { Val = "28" };

            runProperties230.Append(runFonts288);
            runProperties230.Append(bold258);
            runProperties230.Append(color138);
            runProperties230.Append(fontSize239);
            Text text129 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text129.Text = " ~ ";

            run230.Append(runProperties230);
            run230.Append(text129);

            DeletedRun deletedRun57 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "72" };

            Run run231 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties231 = new RunProperties();
            RunFonts runFonts289 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold259 = new Bold();
            Color color139 = new Color() { Val = "0000FF" };
            FontSize fontSize240 = new FontSize() { Val = "28" };

            runProperties231.Append(runFonts289);
            runProperties231.Append(bold259);
            runProperties231.Append(color139);
            runProperties231.Append(fontSize240);
            DeletedText deletedText99 = new DeletedText();
            deletedText99.Text = "[";

            run231.Append(runProperties231);
            run231.Append(deletedText99);

            Run run232 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties232 = new RunProperties();
            RunFonts runFonts290 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold260 = new Bold();
            Color color140 = new Color() { Val = "0000FF" };
            FontSize fontSize241 = new FontSize() { Val = "28" };

            runProperties232.Append(runFonts290);
            runProperties232.Append(bold260);
            runProperties232.Append(color140);
            runProperties232.Append(fontSize241);
            DeletedText deletedText100 = new DeletedText();
            deletedText100.Text = "查程";

            run232.Append(runProperties232);
            run232.Append(deletedText100);

            Run run233 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties233 = new RunProperties();
            RunFonts runFonts291 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold261 = new Bold();
            Color color141 = new Color() { Val = "0000FF" };
            FontSize fontSize242 = new FontSize() { Val = "28" };

            runProperties233.Append(runFonts291);
            runProperties233.Append(bold261);
            runProperties233.Append(color141);
            runProperties233.Append(fontSize242);
            DeletedText deletedText101 = new DeletedText();
            deletedText101.Text = "].[";

            run233.Append(runProperties233);
            run233.Append(deletedText101);

            deletedRun57.Append(run231);
            deletedRun57.Append(run232);
            deletedRun57.Append(run233);

            Run run234 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties234 = new RunProperties();
            RunFonts runFonts292 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold262 = new Bold();
            Color color142 = new Color() { Val = "0000FF" };
            FontSize fontSize243 = new FontSize() { Val = "28" };

            runProperties234.Append(runFonts292);
            runProperties234.Append(bold262);
            runProperties234.Append(color142);
            runProperties234.Append(fontSize243);
            Text text130 = new Text();
            text130.Text = "查核範圍迄日";

            run234.Append(runProperties234);
            run234.Append(text130);

            DeletedRun deletedRun58 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "73" };

            Run run235 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties235 = new RunProperties();
            RunFonts runFonts293 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold263 = new Bold();
            Color color143 = new Color() { Val = "0000FF" };
            FontSize fontSize244 = new FontSize() { Val = "28" };

            runProperties235.Append(runFonts293);
            runProperties235.Append(bold263);
            runProperties235.Append(color143);
            runProperties235.Append(fontSize244);
            DeletedText deletedText102 = new DeletedText();
            deletedText102.Text = "]";

            run235.Append(runProperties235);
            run235.Append(deletedText102);

            deletedRun58.Append(run235);

            paragraph58.Append(paragraphProperties58);
            paragraph58.Append(deletedRun55);
            paragraph58.Append(run228);
            paragraph58.Append(deletedRun56);
            paragraph58.Append(run230);
            paragraph58.Append(deletedRun57);
            paragraph58.Append(run234);
            paragraph58.Append(deletedRun58);

            tableCell36.Append(tableCellProperties36);
            tableCell36.Append(paragraph58);

            tableRow16.Append(tableCell35);
            tableRow16.Append(tableCell36);

            TableRow tableRow17 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "3D7D50D9", TextId = "77777777" };

            TableCell tableCell37 = new TableCell();

            TableCellProperties tableCellProperties37 = new TableCellProperties();
            TableCellWidth tableCellWidth37 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties37.Append(tableCellWidth37);

            Paragraph paragraph59 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "02D4EA82", TextId = "77777777" };

            ParagraphProperties paragraphProperties59 = new ParagraphProperties();
            SnapToGrid snapToGrid51 = new SnapToGrid() { Val = false };
            Indentation indentation32 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification41 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties59 = new ParagraphMarkRunProperties();
            RunFonts runFonts294 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold264 = new Bold();
            FontSize fontSize245 = new FontSize() { Val = "28" };
            Languages languages29 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties59.Append(runFonts294);
            paragraphMarkRunProperties59.Append(bold264);
            paragraphMarkRunProperties59.Append(fontSize245);
            paragraphMarkRunProperties59.Append(languages29);

            paragraphProperties59.Append(snapToGrid51);
            paragraphProperties59.Append(indentation32);
            paragraphProperties59.Append(justification41);
            paragraphProperties59.Append(paragraphMarkRunProperties59);

            Run run236 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties236 = new RunProperties();
            RunFonts runFonts295 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold265 = new Bold();
            FontSize fontSize246 = new FontSize() { Val = "28" };
            Languages languages30 = new Languages() { EastAsia = "zh-HK" };

            runProperties236.Append(runFonts295);
            runProperties236.Append(bold265);
            runProperties236.Append(fontSize246);
            runProperties236.Append(languages30);
            Text text131 = new Text();
            text131.Text = "Auditor in-charge:";

            run236.Append(runProperties236);
            run236.Append(text131);

            paragraph59.Append(paragraphProperties59);
            paragraph59.Append(run236);

            tableCell37.Append(tableCellProperties37);
            tableCell37.Append(paragraph59);

            TableCell tableCell38 = new TableCell();

            TableCellProperties tableCellProperties38 = new TableCellProperties();
            TableCellWidth tableCellWidth38 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties38.Append(tableCellWidth38);

            Paragraph paragraph60 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "3D35369F", TextId = "77777777" };

            ParagraphProperties paragraphProperties60 = new ParagraphProperties();
            SnapToGrid snapToGrid52 = new SnapToGrid() { Val = false };
            Indentation indentation33 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification42 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            RunFonts runFonts296 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold266 = new Bold();
            Color color144 = new Color() { Val = "0000FF" };
            FontSize fontSize247 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties60.Append(runFonts296);
            paragraphMarkRunProperties60.Append(bold266);
            paragraphMarkRunProperties60.Append(color144);
            paragraphMarkRunProperties60.Append(fontSize247);

            paragraphProperties60.Append(snapToGrid52);
            paragraphProperties60.Append(indentation33);
            paragraphProperties60.Append(justification42);
            paragraphProperties60.Append(paragraphMarkRunProperties60);

            DeletedRun deletedRun59 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "74" };

            Run run237 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties237 = new RunProperties();
            RunFonts runFonts297 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold267 = new Bold();
            Color color145 = new Color() { Val = "0000FF" };
            FontSize fontSize248 = new FontSize() { Val = "28" };

            runProperties237.Append(runFonts297);
            runProperties237.Append(bold267);
            runProperties237.Append(color145);
            runProperties237.Append(fontSize248);
            DeletedText deletedText103 = new DeletedText();
            deletedText103.Text = "[";

            run237.Append(runProperties237);
            run237.Append(deletedText103);

            Run run238 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties238 = new RunProperties();
            RunFonts runFonts298 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold268 = new Bold();
            Color color146 = new Color() { Val = "0000FF" };
            FontSize fontSize249 = new FontSize() { Val = "28" };

            runProperties238.Append(runFonts298);
            runProperties238.Append(bold268);
            runProperties238.Append(color146);
            runProperties238.Append(fontSize249);
            DeletedText deletedText104 = new DeletedText();
            deletedText104.Text = "查程";

            run238.Append(runProperties238);
            run238.Append(deletedText104);

            Run run239 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties239 = new RunProperties();
            RunFonts runFonts299 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold269 = new Bold();
            Color color147 = new Color() { Val = "0000FF" };
            FontSize fontSize250 = new FontSize() { Val = "28" };

            runProperties239.Append(runFonts299);
            runProperties239.Append(bold269);
            runProperties239.Append(color147);
            runProperties239.Append(fontSize250);
            DeletedText deletedText105 = new DeletedText();
            deletedText105.Text = "].[";

            run239.Append(runProperties239);
            run239.Append(deletedText105);

            deletedRun59.Append(run237);
            deletedRun59.Append(run238);
            deletedRun59.Append(run239);

            Run run240 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties240 = new RunProperties();
            RunFonts runFonts300 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold270 = new Bold();
            Color color148 = new Color() { Val = "0000FF" };
            FontSize fontSize251 = new FontSize() { Val = "28" };

            runProperties240.Append(runFonts300);
            runProperties240.Append(bold270);
            runProperties240.Append(color148);
            runProperties240.Append(fontSize251);
            Text text132 = new Text();
            text132.Text = dt.Rows[0]["leader"].ToString();

            run240.Append(runProperties240);
            run240.Append(text132);

            DeletedRun deletedRun60 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "75" };

            Run run241 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "0096615A" };

            RunProperties runProperties241 = new RunProperties();
            RunFonts runFonts301 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold271 = new Bold();
            Color color149 = new Color() { Val = "0000FF" };
            FontSize fontSize252 = new FontSize() { Val = "28" };

            runProperties241.Append(runFonts301);
            runProperties241.Append(bold271);
            runProperties241.Append(color149);
            runProperties241.Append(fontSize252);
            DeletedText deletedText106 = new DeletedText();
            deletedText106.Text = "_Eng";

            run241.Append(runProperties241);
            run241.Append(deletedText106);

            deletedRun60.Append(run241);

            DeletedRun deletedRun61 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "76" };

            Run run242 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties242 = new RunProperties();
            RunFonts runFonts302 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold272 = new Bold();
            Color color150 = new Color() { Val = "0000FF" };
            FontSize fontSize253 = new FontSize() { Val = "28" };

            runProperties242.Append(runFonts302);
            runProperties242.Append(bold272);
            runProperties242.Append(color150);
            runProperties242.Append(fontSize253);
            DeletedText deletedText107 = new DeletedText();
            deletedText107.Text = "]";

            run242.Append(runProperties242);
            run242.Append(deletedText107);

            deletedRun61.Append(run242);

            paragraph60.Append(paragraphProperties60);
            paragraph60.Append(deletedRun59);
            paragraph60.Append(run240);
            paragraph60.Append(deletedRun60);
            paragraph60.Append(deletedRun61);

            tableCell38.Append(tableCellProperties38);
            tableCell38.Append(paragraph60);

            tableRow17.Append(tableCell37);
            tableRow17.Append(tableCell38);

            TableRow tableRow18 = new TableRow() { RsidTableRowMarkRevision = "001752F7", RsidTableRowAddition = "001752F7", RsidTableRowProperties = "0096027D", ParagraphId = "19ECF03D", TextId = "77777777" };

            TableCell tableCell39 = new TableCell();

            TableCellProperties tableCellProperties39 = new TableCellProperties();
            TableCellWidth tableCellWidth39 = new TableCellWidth() { Width = "2410", Type = TableWidthUnitValues.Dxa };

            tableCellProperties39.Append(tableCellWidth39);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "0B937292", TextId = "77777777" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();
            SnapToGrid snapToGrid53 = new SnapToGrid() { Val = false };
            Indentation indentation34 = new Indentation() { Start = "-2", StartCharacters = -1, FirstLine = "1" };
            Justification justification43 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            RunFonts runFonts303 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold273 = new Bold();
            FontSize fontSize254 = new FontSize() { Val = "28" };
            Languages languages31 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties61.Append(runFonts303);
            paragraphMarkRunProperties61.Append(bold273);
            paragraphMarkRunProperties61.Append(fontSize254);
            paragraphMarkRunProperties61.Append(languages31);

            paragraphProperties61.Append(snapToGrid53);
            paragraphProperties61.Append(indentation34);
            paragraphProperties61.Append(justification43);
            paragraphProperties61.Append(paragraphMarkRunProperties61);

            Run run243 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties243 = new RunProperties();
            RunFonts runFonts304 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold274 = new Bold();
            FontSize fontSize255 = new FontSize() { Val = "28" };
            Languages languages32 = new Languages() { EastAsia = "zh-HK" };

            runProperties243.Append(runFonts304);
            runProperties243.Append(bold274);
            runProperties243.Append(fontSize255);
            runProperties243.Append(languages32);
            Text text133 = new Text();
            text133.Text = "Auditor:";

            run243.Append(runProperties243);
            run243.Append(text133);

            paragraph61.Append(paragraphProperties61);
            paragraph61.Append(run243);

            tableCell39.Append(tableCellProperties39);
            tableCell39.Append(paragraph61);

            TableCell tableCell40 = new TableCell();

            TableCellProperties tableCellProperties40 = new TableCellProperties();
            TableCellWidth tableCellWidth40 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties40.Append(tableCellWidth40);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphMarkRevision = "001752F7", RsidParagraphAddition = "001752F7", RsidParagraphProperties = "0096027D", RsidRunAdditionDefault = "001752F7", ParagraphId = "329F076F", TextId = "77777777" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();
            SnapToGrid snapToGrid54 = new SnapToGrid() { Val = false };
            Indentation indentation35 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification44 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties62 = new ParagraphMarkRunProperties();
            RunFonts runFonts305 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold275 = new Bold();
            Color color151 = new Color() { Val = "0000FF" };
            FontSize fontSize256 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties62.Append(runFonts305);
            paragraphMarkRunProperties62.Append(bold275);
            paragraphMarkRunProperties62.Append(color151);
            paragraphMarkRunProperties62.Append(fontSize256);

            paragraphProperties62.Append(snapToGrid54);
            paragraphProperties62.Append(indentation35);
            paragraphProperties62.Append(justification44);
            paragraphProperties62.Append(paragraphMarkRunProperties62);

            DeletedRun deletedRun62 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "77" };

            Run run244 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties244 = new RunProperties();
            RunFonts runFonts306 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold276 = new Bold();
            Color color152 = new Color() { Val = "0000FF" };
            FontSize fontSize257 = new FontSize() { Val = "28" };

            runProperties244.Append(runFonts306);
            runProperties244.Append(bold276);
            runProperties244.Append(color152);
            runProperties244.Append(fontSize257);
            DeletedText deletedText108 = new DeletedText();
            deletedText108.Text = "UNION([";

            run244.Append(runProperties244);
            run244.Append(deletedText108);

            Run run245 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties245 = new RunProperties();
            RunFonts runFonts307 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold277 = new Bold();
            Color color153 = new Color() { Val = "0000FF" };
            FontSize fontSize258 = new FontSize() { Val = "28" };

            runProperties245.Append(runFonts307);
            runProperties245.Append(bold277);
            runProperties245.Append(color153);
            runProperties245.Append(fontSize258);
            DeletedText deletedText109 = new DeletedText();
            deletedText109.Text = "查程工作分配";

            run245.Append(runProperties245);
            run245.Append(deletedText109);

            Run run246 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties246 = new RunProperties();
            RunFonts runFonts308 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold278 = new Bold();
            Color color154 = new Color() { Val = "0000FF" };
            FontSize fontSize259 = new FontSize() { Val = "28" };

            runProperties246.Append(runFonts308);
            runProperties246.Append(bold278);
            runProperties246.Append(color154);
            runProperties246.Append(fontSize259);
            DeletedText deletedText110 = new DeletedText();
            deletedText110.Text = "].[";

            run246.Append(runProperties246);
            run246.Append(deletedText110);

            deletedRun62.Append(run244);
            deletedRun62.Append(run245);
            deletedRun62.Append(run246);

            Run run247 = new Run() { RsidRunProperties = "001752F7" };

            RunProperties runProperties247 = new RunProperties();
            RunFonts runFonts309 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold279 = new Bold();
            Color color155 = new Color() { Val = "0000FF" };
            FontSize fontSize260 = new FontSize() { Val = "28" };

            runProperties247.Append(runFonts309);
            runProperties247.Append(bold279);
            runProperties247.Append(color155);
            runProperties247.Append(fontSize260);
            Text text134 = new Text();
            text134.Text = dt.Rows[0]["Member"].ToString();

            run247.Append(runProperties247);
            run247.Append(text134);

            DeletedRun deletedRun63 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:52:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "78" };

            Run run248 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "0096615A" };

            RunProperties runProperties248 = new RunProperties();
            RunFonts runFonts310 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold280 = new Bold();
            Color color156 = new Color() { Val = "0000FF" };
            FontSize fontSize261 = new FontSize() { Val = "28" };

            runProperties248.Append(runFonts310);
            runProperties248.Append(bold280);
            runProperties248.Append(color156);
            runProperties248.Append(fontSize261);
            DeletedText deletedText111 = new DeletedText();
            deletedText111.Text = "_Eng";

            run248.Append(runProperties248);
            run248.Append(deletedText111);

            deletedRun63.Append(run248);

            DeletedRun deletedRun64 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:05:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "79" };

            Run run249 = new Run() { RsidRunProperties = "001752F7", RsidRunDeletion = "001C7370" };

            RunProperties runProperties249 = new RunProperties();
            RunFonts runFonts311 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold281 = new Bold();
            Color color157 = new Color() { Val = "0000FF" };
            FontSize fontSize262 = new FontSize() { Val = "28" };

            runProperties249.Append(runFonts311);
            runProperties249.Append(bold281);
            runProperties249.Append(color157);
            runProperties249.Append(fontSize262);
            DeletedText deletedText112 = new DeletedText();
            deletedText112.Text = "])";

            run249.Append(runProperties249);
            run249.Append(deletedText112);

            deletedRun64.Append(run249);

            paragraph62.Append(paragraphProperties62);
            paragraph62.Append(deletedRun62);
            paragraph62.Append(run247);
            paragraph62.Append(deletedRun63);
            paragraph62.Append(deletedRun64);

            tableCell40.Append(tableCellProperties40);
            tableCell40.Append(paragraph62);

            tableRow18.Append(tableCell39);
            tableRow18.Append(tableCell40);

            table3.Append(tableProperties3);
            table3.Append(tableGrid3);
            table3.Append(tableRow12);
            table3.Append(tableRow13);
            table3.Append(tableRow14);
            table3.Append(tableRow15);
            table3.Append(tableRow16);
            table3.Append(tableRow17);
            table3.Append(tableRow18);

            Paragraph paragraph63 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00200048", RsidRunAdditionDefault = "00200048", ParagraphId = "31C90F81", TextId = "77777777" };

            ParagraphProperties paragraphProperties63 = new ParagraphProperties();
            SnapToGrid snapToGrid55 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties63 = new ParagraphMarkRunProperties();
            RunFonts runFonts312 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties63.Append(runFonts312);

            paragraphProperties63.Append(snapToGrid55);
            paragraphProperties63.Append(spacingBetweenLines9);
            paragraphProperties63.Append(paragraphMarkRunProperties63);

            paragraph63.Append(paragraphProperties63);

            Paragraph paragraph64 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "009B7102", RsidRunAdditionDefault = "00200048", ParagraphId = "09649C50", TextId = "77777777" };

            ParagraphProperties paragraphProperties64 = new ParagraphProperties();

            NumberingProperties numberingProperties11 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference11 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId11 = new NumberingId() { Val = 10 };

            numberingProperties11.Append(numberingLevelReference11);
            numberingProperties11.Append(numberingId11);
            SnapToGrid snapToGrid56 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties64 = new ParagraphMarkRunProperties();
            RunFonts runFonts313 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize263 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties64.Append(runFonts313);
            paragraphMarkRunProperties64.Append(fontSize263);

            paragraphProperties64.Append(numberingProperties11);
            paragraphProperties64.Append(snapToGrid56);
            paragraphProperties64.Append(spacingBetweenLines10);
            paragraphProperties64.Append(paragraphMarkRunProperties64);

            Run run250 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties250 = new RunProperties();
            RunFonts runFonts314 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize264 = new FontSize() { Val = "28" };

            runProperties250.Append(runFonts314);
            runProperties250.Append(fontSize264);
            Text text135 = new Text();
            text135.Text = "「";

            run250.Append(runProperties250);
            run250.Append(text135);

            Run run251 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "000B740F" };

            RunProperties runProperties251 = new RunProperties();
            RunFonts runFonts315 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color158 = new Color() { Val = "0000FF" };
            FontSize fontSize265 = new FontSize() { Val = "28" };

            runProperties251.Append(runFonts315);
            runProperties251.Append(color158);
            runProperties251.Append(fontSize265);
            Text text136 = new Text();
            text136.Text = "XXXXXXXXX";

            run251.Append(runProperties251);
            run251.Append(text136);

            Run run252 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties252 = new RunProperties();
            RunFonts runFonts316 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize266 = new FontSize() { Val = "28" };

            runProperties252.Append(runFonts316);
            runProperties252.Append(fontSize266);
            Text text137 = new Text();
            text137.Text = "」";

            run252.Append(runProperties252);
            run252.Append(text137);

            Run run253 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "009B7102" };

            RunProperties runProperties253 = new RunProperties();
            RunFonts runFonts317 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize267 = new FontSize() { Val = "28" };

            runProperties253.Append(runFonts317);
            runProperties253.Append(fontSize267);
            Text text138 = new Text();
            text138.Text = "R";

            run253.Append(runProperties253);
            run253.Append(text138);

            Run run254 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "009B7102" };

            RunProperties runProperties254 = new RunProperties();
            RunFonts runFonts318 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize268 = new FontSize() { Val = "28" };

            runProperties254.Append(runFonts318);
            runProperties254.Append(fontSize268);
            Text text139 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text139.Text = "isk and ";

            run254.Append(runProperties254);
            run254.Append(text139);

            Run run255 = new Run() { RsidRunProperties = "009B7102", RsidRunAddition = "009B7102" };

            RunProperties runProperties255 = new RunProperties();
            RunFonts runFonts319 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize269 = new FontSize() { Val = "28" };

            runProperties255.Append(runFonts319);
            runProperties255.Append(fontSize269);
            Text text140 = new Text();
            text140.Text = "Observations from Process Understanding";

            run255.Append(runProperties255);
            run255.Append(text140);

            Run run256 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "009B7102" };

            RunProperties runProperties256 = new RunProperties();
            RunFonts runFonts320 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize270 = new FontSize() { Val = "28" };

            runProperties256.Append(runFonts320);
            runProperties256.Append(fontSize270);
            Text text141 = new Text();
            text141.Text = ":";

            run256.Append(runProperties256);
            run256.Append(text141);

            paragraph64.Append(paragraphProperties64);
            paragraph64.Append(run250);
            paragraph64.Append(run251);
            paragraph64.Append(run252);
            paragraph64.Append(run253);
            paragraph64.Append(run254);
            paragraph64.Append(run255);
            paragraph64.Append(run256);

            Paragraph paragraph65 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "001C40CD", RsidRunAdditionDefault = "009B7102", ParagraphId = "73CD515F", TextId = "77777777" };

            ParagraphProperties paragraphProperties65 = new ParagraphProperties();

            NumberingProperties numberingProperties12 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference12 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId12 = new NumberingId() { Val = 17 };

            numberingProperties12.Append(numberingLevelReference12);
            numberingProperties12.Append(numberingId12);

            Tabs tabs5 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Left, Position = 1134 };

            tabs5.Append(tabStop5);
            SnapToGrid snapToGrid57 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties65 = new ParagraphMarkRunProperties();
            RunFonts runFonts321 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize271 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties65.Append(runFonts321);
            paragraphMarkRunProperties65.Append(fontSize271);

            paragraphProperties65.Append(numberingProperties12);
            paragraphProperties65.Append(tabs5);
            paragraphProperties65.Append(snapToGrid57);
            paragraphProperties65.Append(spacingBetweenLines11);
            paragraphProperties65.Append(paragraphMarkRunProperties65);

            Run run257 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties257 = new RunProperties();
            RunFonts runFonts322 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize272 = new FontSize() { Val = "28" };

            runProperties257.Append(runFonts322);
            runProperties257.Append(fontSize272);
            Text text142 = new Text();
            text142.Text = "Key risk";

            run257.Append(runProperties257);
            run257.Append(text142);

            Run run258 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "007E743C" };

            RunProperties runProperties258 = new RunProperties();
            RunFonts runFonts323 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize273 = new FontSize() { Val = "28" };

            runProperties258.Append(runFonts323);
            runProperties258.Append(fontSize273);
            Text text143 = new Text();
            text143.Text = "s";

            run258.Append(runProperties258);
            run258.Append(text143);

            Run run259 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties259 = new RunProperties();
            RunFonts runFonts324 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize274 = new FontSize() { Val = "28" };

            runProperties259.Append(runFonts324);
            runProperties259.Append(fontSize274);
            Text text144 = new Text();
            text144.Text = ":";

            run259.Append(runProperties259);
            run259.Append(text144);

            paragraph65.Append(paragraphProperties65);
            paragraph65.Append(run257);
            paragraph65.Append(run258);
            paragraph65.Append(run259);

            Paragraph paragraph66 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "001C40CD", RsidRunAdditionDefault = "003F405A", ParagraphId = "26B79F13", TextId = "77777777" };

            ParagraphProperties paragraphProperties66 = new ParagraphProperties();

            NumberingProperties numberingProperties13 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference13 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId13 = new NumberingId() { Val = 17 };

            numberingProperties13.Append(numberingLevelReference13);
            numberingProperties13.Append(numberingId13);

            Tabs tabs6 = new Tabs();
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Number, Position = 1134 };

            tabs6.Append(tabStop6);
            SnapToGrid snapToGrid58 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties66 = new ParagraphMarkRunProperties();
            RunFonts runFonts325 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize275 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties66.Append(runFonts325);
            paragraphMarkRunProperties66.Append(fontSize275);

            paragraphProperties66.Append(numberingProperties13);
            paragraphProperties66.Append(tabs6);
            paragraphProperties66.Append(snapToGrid58);
            paragraphProperties66.Append(spacingBetweenLines12);
            paragraphProperties66.Append(paragraphMarkRunProperties66);

            Run run260 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties260 = new RunProperties();
            RunFonts runFonts326 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize276 = new FontSize() { Val = "28" };

            runProperties260.Append(runFonts326);
            runProperties260.Append(fontSize276);
            Text text145 = new Text();
            text145.Text = "C";

            run260.Append(runProperties260);
            run260.Append(text145);

            Run run261 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties261 = new RunProperties();
            RunFonts runFonts327 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize277 = new FontSize() { Val = "28" };

            runProperties261.Append(runFonts327);
            runProperties261.Append(fontSize277);
            Text text146 = new Text();
            text146.Text = "ontrol process";

            run261.Append(runProperties261);
            run261.Append(text146);

            Run run262 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "007E743C" };

            RunProperties runProperties262 = new RunProperties();
            RunFonts runFonts328 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize278 = new FontSize() { Val = "28" };

            runProperties262.Append(runFonts328);
            runProperties262.Append(fontSize278);
            Text text147 = new Text();
            text147.Text = "es";

            run262.Append(runProperties262);
            run262.Append(text147);

            Run run263 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "00010E46" };

            RunProperties runProperties263 = new RunProperties();
            RunFonts runFonts329 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize279 = new FontSize() { Val = "28" };

            runProperties263.Append(runFonts329);
            runProperties263.Append(fontSize279);
            Text text148 = new Text();
            text148.Text = ":";

            run263.Append(runProperties263);
            run263.Append(text148);

            paragraph66.Append(paragraphProperties66);
            paragraph66.Append(run260);
            paragraph66.Append(run261);
            paragraph66.Append(run262);
            paragraph66.Append(run263);

            Paragraph paragraph67 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "001C40CD", RsidRunAdditionDefault = "007E743C", ParagraphId = "6D5E7597", TextId = "77777777" };

            ParagraphProperties paragraphProperties67 = new ParagraphProperties();

            NumberingProperties numberingProperties14 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference14 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId14 = new NumberingId() { Val = 17 };

            numberingProperties14.Append(numberingLevelReference14);
            numberingProperties14.Append(numberingId14);

            Tabs tabs7 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 1134 };

            tabs7.Append(tabStop7);
            SnapToGrid snapToGrid59 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties67 = new ParagraphMarkRunProperties();
            RunFonts runFonts330 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize280 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties67.Append(runFonts330);
            paragraphMarkRunProperties67.Append(fontSize280);

            paragraphProperties67.Append(numberingProperties14);
            paragraphProperties67.Append(tabs7);
            paragraphProperties67.Append(snapToGrid59);
            paragraphProperties67.Append(spacingBetweenLines13);
            paragraphProperties67.Append(paragraphMarkRunProperties67);

            Run run264 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties264 = new RunProperties();
            RunFonts runFonts331 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize281 = new FontSize() { Val = "28" };

            runProperties264.Append(runFonts331);
            runProperties264.Append(fontSize281);
            Text text149 = new Text();
            text149.Text = "G";

            run264.Append(runProperties264);
            run264.Append(text149);

            Run run265 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties265 = new RunProperties();
            RunFonts runFonts332 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize282 = new FontSize() { Val = "28" };

            runProperties265.Append(runFonts332);
            runProperties265.Append(fontSize282);
            Text text150 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text150.Text = "AP (or any ";

            run265.Append(runProperties265);
            run265.Append(text150);

            Run run266 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties266 = new RunProperties();
            RunFonts runFonts333 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize283 = new FontSize() { Val = "28" };

            runProperties266.Append(runFonts333);
            runProperties266.Append(fontSize283);
            Text text151 = new Text();
            text151.Text = "m";

            run266.Append(runProperties266);
            run266.Append(text151);

            Run run267 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties267 = new RunProperties();
            RunFonts runFonts334 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize284 = new FontSize() { Val = "28" };

            runProperties267.Append(runFonts334);
            runProperties267.Append(fontSize284);
            Text text152 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text152.Text = "ajor deficiencies ";

            run267.Append(runProperties267);
            run267.Append(text152);

            Run run268 = new Run() { RsidRunAddition = "00BA581D" };

            RunProperties runProperties268 = new RunProperties();
            RunFonts runFonts335 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize285 = new FontSize() { Val = "28" };

            runProperties268.Append(runFonts335);
            runProperties268.Append(fontSize285);
            Text text153 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text153.Text = "have ";

            run268.Append(runProperties268);
            run268.Append(text153);

            Run run269 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties269 = new RunProperties();
            RunFonts runFonts336 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize286 = new FontSize() { Val = "28" };

            runProperties269.Append(runFonts336);
            runProperties269.Append(fontSize286);
            Text text154 = new Text();
            text154.Text = "been identified):";

            run269.Append(runProperties269);
            run269.Append(text154);

            paragraph67.Append(paragraphProperties67);
            paragraph67.Append(run264);
            paragraph67.Append(run265);
            paragraph67.Append(run266);
            paragraph67.Append(run267);
            paragraph67.Append(run268);
            paragraph67.Append(run269);

            Paragraph paragraph68 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "001C40CD", RsidRunAdditionDefault = "001C40CD", ParagraphId = "5BD20BF4", TextId = "77777777" };

            ParagraphProperties paragraphProperties68 = new ParagraphProperties();

            NumberingProperties numberingProperties15 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference15 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId15 = new NumberingId() { Val = 17 };

            numberingProperties15.Append(numberingLevelReference15);
            numberingProperties15.Append(numberingId15);

            Tabs tabs8 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 1134 };

            tabs8.Append(tabStop8);
            SnapToGrid snapToGrid60 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties68 = new ParagraphMarkRunProperties();
            RunFonts runFonts337 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize287 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties68.Append(runFonts337);
            paragraphMarkRunProperties68.Append(fontSize287);

            paragraphProperties68.Append(numberingProperties15);
            paragraphProperties68.Append(tabs8);
            paragraphProperties68.Append(snapToGrid60);
            paragraphProperties68.Append(spacingBetweenLines14);
            paragraphProperties68.Append(paragraphMarkRunProperties68);

            Run run270 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties270 = new RunProperties();
            RunFonts runFonts338 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize288 = new FontSize() { Val = "28" };

            runProperties270.Append(runFonts338);
            runProperties270.Append(fontSize288);
            Text text155 = new Text();
            text155.Text = "Countermeasures:";

            run270.Append(runProperties270);
            run270.Append(text155);

            paragraph68.Append(paragraphProperties68);
            paragraph68.Append(run270);

            Paragraph paragraph69 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "001C40CD", RsidParagraphProperties = "001C40CD", RsidRunAdditionDefault = "001C40CD", ParagraphId = "24217146", TextId = "77777777" };

            ParagraphProperties paragraphProperties69 = new ParagraphProperties();
            SnapToGrid snapToGrid61 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation36 = new Indentation() { Start = "1680" };

            ParagraphMarkRunProperties paragraphMarkRunProperties69 = new ParagraphMarkRunProperties();
            RunFonts runFonts339 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize289 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties69.Append(runFonts339);
            paragraphMarkRunProperties69.Append(fontSize289);

            paragraphProperties69.Append(snapToGrid61);
            paragraphProperties69.Append(spacingBetweenLines15);
            paragraphProperties69.Append(indentation36);
            paragraphProperties69.Append(paragraphMarkRunProperties69);

            paragraph69.Append(paragraphProperties69);

            Paragraph paragraph70 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "005D00FB", RsidRunAdditionDefault = "005D00FB", ParagraphId = "320357E8", TextId = "421503D3" };

            ParagraphProperties paragraphProperties70 = new ParagraphProperties();

            NumberingProperties numberingProperties16 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference16 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId16 = new NumberingId() { Val = 10 };

            numberingProperties16.Append(numberingLevelReference16);
            numberingProperties16.Append(numberingId16);
            SnapToGrid snapToGrid62 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties70 = new ParagraphMarkRunProperties();
            RunFonts runFonts340 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize290 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties70.Append(runFonts340);
            paragraphMarkRunProperties70.Append(fontSize290);

            paragraphProperties70.Append(numberingProperties16);
            paragraphProperties70.Append(snapToGrid62);
            paragraphProperties70.Append(spacingBetweenLines16);
            paragraphProperties70.Append(paragraphMarkRunProperties70);

            Run run271 = new Run() { RsidRunProperties = "005D00FB" };

            RunProperties runProperties271 = new RunProperties();
            RunFonts runFonts341 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize291 = new FontSize() { Val = "28" };

            runProperties271.Append(runFonts341);
            runProperties271.Append(fontSize291);
            Text text156 = new Text();
            text156.Text = "Objective";

            run271.Append(runProperties271);
            run271.Append(text156);

            Run run272 = new Run();

            RunProperties runProperties272 = new RunProperties();
            RunFonts runFonts342 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize292 = new FontSize() { Val = "28" };

            runProperties272.Append(runFonts342);
            runProperties272.Append(fontSize292);
            Text text157 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text157.Text = " of ";

            run272.Append(runProperties272);
            run272.Append(text157);

            Run run273 = new Run() { RsidRunAddition = "004C41E7" };

            RunProperties runProperties273 = new RunProperties();
            RunFonts runFonts343 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize293 = new FontSize() { Val = "28" };

            runProperties273.Append(runFonts343);
            runProperties273.Append(fontSize293);
            Text text158 = new Text();
            text158.Text = "E";

            run273.Append(runProperties273);
            run273.Append(text158);

            Run run274 = new Run();

            RunProperties runProperties274 = new RunProperties();
            RunFonts runFonts344 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize294 = new FontSize() { Val = "28" };

            runProperties274.Append(runFonts344);
            runProperties274.Append(fontSize294);
            Text text159 = new Text();
            text159.Text = "ngagement:";

            run274.Append(runProperties274);
            run274.Append(text159);

            paragraph70.Append(paragraphProperties70);
            paragraph70.Append(run271);
            paragraph70.Append(run272);
            paragraph70.Append(run273);
            paragraph70.Append(run274);

            Paragraph paragraph71 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "005D00FB", ParagraphId = "45503214", TextId = "77777777" };

            ParagraphProperties paragraphProperties71 = new ParagraphProperties();

            NumberingProperties numberingProperties17 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference17 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId17 = new NumberingId() { Val = 10 };

            numberingProperties17.Append(numberingLevelReference17);
            numberingProperties17.Append(numberingId17);
            SnapToGrid snapToGrid63 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation37 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties71 = new ParagraphMarkRunProperties();
            RunFonts runFonts345 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize295 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties71.Append(runFonts345);
            paragraphMarkRunProperties71.Append(fontSize295);

            paragraphProperties71.Append(numberingProperties17);
            paragraphProperties71.Append(snapToGrid63);
            paragraphProperties71.Append(spacingBetweenLines17);
            paragraphProperties71.Append(indentation37);
            paragraphProperties71.Append(paragraphMarkRunProperties71);

            Run run275 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties275 = new RunProperties();
            RunFonts runFonts346 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize296 = new FontSize() { Val = "28" };

            runProperties275.Append(runFonts346);
            runProperties275.Append(fontSize296);
            Text text160 = new Text();
            text160.Text = "Audit Focus:";

            run275.Append(runProperties275);
            run275.Append(text160);

            paragraph71.Append(paragraphProperties71);
            paragraph71.Append(run275);

            Paragraph paragraph72 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00567A5E", ParagraphId = "309DF817", TextId = "45CC4768" };

            ParagraphProperties paragraphProperties72 = new ParagraphProperties();

            NumberingProperties numberingProperties18 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference18 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId18 = new NumberingId() { Val = 10 };

            numberingProperties18.Append(numberingLevelReference18);
            numberingProperties18.Append(numberingId18);
            SnapToGrid snapToGrid64 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation38 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties72 = new ParagraphMarkRunProperties();
            RunFonts runFonts347 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize297 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties72.Append(runFonts347);
            paragraphMarkRunProperties72.Append(fontSize297);

            paragraphProperties72.Append(numberingProperties18);
            paragraphProperties72.Append(snapToGrid64);
            paragraphProperties72.Append(spacingBetweenLines18);
            paragraphProperties72.Append(indentation38);
            paragraphProperties72.Append(paragraphMarkRunProperties72);

            Run run276 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties276 = new RunProperties();
            RunFonts runFonts348 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize298 = new FontSize() { Val = "28" };

            runProperties276.Append(runFonts348);
            runProperties276.Append(fontSize298);
            Text text161 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text161.Text = "Job ";

            run276.Append(runProperties276);
            run276.Append(text161);

            Run run277 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "00153D64" };

            RunProperties runProperties277 = new RunProperties();
            RunFonts runFonts349 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize299 = new FontSize() { Val = "28" };

            runProperties277.Append(runFonts349);
            runProperties277.Append(fontSize299);
            Text text162 = new Text();
            text162.Text = "Assignments";

            run277.Append(runProperties277);
            run277.Append(text162);

            Run run278 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties278 = new RunProperties();
            RunFonts runFonts350 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize300 = new FontSize() { Val = "28" };

            runProperties278.Append(runFonts350);
            runProperties278.Append(fontSize300);
            Text text163 = new Text();
            text163.Text = ":";

            run278.Append(runProperties278);
            run278.Append(text163);

            paragraph72.Append(paragraphProperties72);
            paragraph72.Append(run276);
            paragraph72.Append(run277);
            paragraph72.Append(run278);

            Paragraph paragraph73 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00567A5E", ParagraphId = "3743677E", TextId = "77777777" };

            ParagraphProperties paragraphProperties73 = new ParagraphProperties();

            NumberingProperties numberingProperties19 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference19 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId19 = new NumberingId() { Val = 10 };

            numberingProperties19.Append(numberingLevelReference19);
            numberingProperties19.Append(numberingId19);
            SnapToGrid snapToGrid65 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation39 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties73 = new ParagraphMarkRunProperties();
            RunFonts runFonts351 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize301 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties73.Append(runFonts351);
            paragraphMarkRunProperties73.Append(fontSize301);

            paragraphProperties73.Append(numberingProperties19);
            paragraphProperties73.Append(snapToGrid65);
            paragraphProperties73.Append(spacingBetweenLines19);
            paragraphProperties73.Append(indentation39);
            paragraphProperties73.Append(paragraphMarkRunProperties73);

            Run run279 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties279 = new RunProperties();
            RunFonts runFonts352 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize302 = new FontSize() { Val = "28" };

            runProperties279.Append(runFonts352);
            runProperties279.Append(fontSize302);
            Text text164 = new Text();
            text164.Text = "Request List";

            run279.Append(runProperties279);
            run279.Append(text164);

            Run run280 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties280 = new RunProperties();
            RunFonts runFonts353 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize303 = new FontSize() { Val = "28" };

            runProperties280.Append(runFonts353);
            runProperties280.Append(fontSize303);
            Text text165 = new Text();
            text165.Text = ":";

            run280.Append(runProperties280);
            run280.Append(text165);

            paragraph73.Append(paragraphProperties73);
            paragraph73.Append(run279);
            paragraph73.Append(run280);

            Paragraph paragraph74 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00214C36", RsidParagraphProperties = "00E6762F", RsidRunAdditionDefault = "00567A5E", ParagraphId = "2FE61E26", TextId = "5F6F3E38" };

            ParagraphProperties paragraphProperties74 = new ParagraphProperties();

            NumberingProperties numberingProperties20 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference20 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId20 = new NumberingId() { Val = 10 };

            numberingProperties20.Append(numberingLevelReference20);
            numberingProperties20.Append(numberingId20);
            SnapToGrid snapToGrid66 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation40 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties74 = new ParagraphMarkRunProperties();
            RunFonts runFonts354 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color159 = new Color() { Val = "000000" };
            FontSize fontSize304 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties74.Append(runFonts354);
            paragraphMarkRunProperties74.Append(color159);
            paragraphMarkRunProperties74.Append(fontSize304);

            paragraphProperties74.Append(numberingProperties20);
            paragraphProperties74.Append(snapToGrid66);
            paragraphProperties74.Append(spacingBetweenLines20);
            paragraphProperties74.Append(indentation40);
            paragraphProperties74.Append(paragraphMarkRunProperties74);

            Run run281 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties281 = new RunProperties();
            RunFonts runFonts355 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color160 = new Color() { Val = "000000" };
            FontSize fontSize305 = new FontSize() { Val = "28" };

            runProperties281.Append(runFonts355);
            runProperties281.Append(color160);
            runProperties281.Append(fontSize305);
            Text text166 = new Text();
            text166.Text = "S";

            run281.Append(runProperties281);
            run281.Append(text166);

            Run run282 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties282 = new RunProperties();
            RunFonts runFonts356 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color161 = new Color() { Val = "000000" };
            FontSize fontSize306 = new FontSize() { Val = "28" };

            runProperties282.Append(runFonts356);
            runProperties282.Append(color161);
            runProperties282.Append(fontSize306);
            Text text167 = new Text();
            text167.Text = "ampling";

            run282.Append(runProperties282);
            run282.Append(text167);

            Run run283 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "00E6762F" };

            RunProperties runProperties283 = new RunProperties();
            RunFonts runFonts357 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color162 = new Color() { Val = "000000" };
            FontSize fontSize307 = new FontSize() { Val = "28" };

            runProperties283.Append(runFonts357);
            runProperties283.Append(color162);
            runProperties283.Append(fontSize307);
            Text text168 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text168.Text = " ";

            run283.Append(runProperties283);
            run283.Append(text168);

            Run run284 = new Run() { RsidRunProperties = "0096027D", RsidRunAddition = "00214C36" };

            RunProperties runProperties284 = new RunProperties();
            RunFonts runFonts358 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color163 = new Color() { Val = "000000" };
            FontSize fontSize308 = new FontSize() { Val = "28" };

            runProperties284.Append(runFonts358);
            runProperties284.Append(color163);
            runProperties284.Append(fontSize308);
            Text text169 = new Text();
            text169.Text = "(";

            run284.Append(runProperties284);
            run284.Append(text169);

            Run run285 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties285 = new RunProperties();
            RunFonts runFonts359 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color164 = new Color() { Val = "000000" };
            FontSize fontSize309 = new FontSize() { Val = "28" };

            runProperties285.Append(runFonts359);
            runProperties285.Append(color164);
            runProperties285.Append(fontSize309);
            Text text170 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text170.Text = "include: method, ";

            run285.Append(runProperties285);
            run285.Append(text170);

            Run run286 = new Run() { RsidRunAddition = "00673A05" };

            RunProperties runProperties286 = new RunProperties();
            RunFonts runFonts360 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color165 = new Color() { Val = "000000" };
            FontSize fontSize310 = new FontSize() { Val = "28" };

            runProperties286.Append(runFonts360);
            runProperties286.Append(color165);
            runProperties286.Append(fontSize310);
            Text text171 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text171.Text = "population ";

            run286.Append(runProperties286);
            run286.Append(text171);

            Run run287 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties287 = new RunProperties();
            RunFonts runFonts361 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color166 = new Color() { Val = "000000" };
            FontSize fontSize311 = new FontSize() { Val = "28" };

            runProperties287.Append(runFonts361);
            runProperties287.Append(color166);
            runProperties287.Append(fontSize311);
            Text text172 = new Text();
            text172.Text = "and number of samples)";

            run287.Append(runProperties287);
            run287.Append(text172);

            Run run288 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties288 = new RunProperties();
            RunFonts runFonts362 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color167 = new Color() { Val = "000000" };
            FontSize fontSize312 = new FontSize() { Val = "28" };

            runProperties288.Append(runFonts362);
            runProperties288.Append(color167);
            runProperties288.Append(fontSize312);
            Text text173 = new Text();
            text173.Text = ":";

            run288.Append(runProperties288);
            run288.Append(text173);

            paragraph74.Append(paragraphProperties74);
            paragraph74.Append(run281);
            paragraph74.Append(run282);
            paragraph74.Append(run283);
            paragraph74.Append(run284);
            paragraph74.Append(run285);
            paragraph74.Append(run286);
            paragraph74.Append(run287);
            paragraph74.Append(run288);

            Paragraph paragraph75 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "004A60EE", ParagraphId = "2A1F5553", TextId = "77777777" };

            ParagraphProperties paragraphProperties75 = new ParagraphProperties();

            NumberingProperties numberingProperties21 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference21 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId21 = new NumberingId() { Val = 10 };

            numberingProperties21.Append(numberingLevelReference21);
            numberingProperties21.Append(numberingId21);
            SnapToGrid snapToGrid67 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation41 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties75 = new ParagraphMarkRunProperties();
            RunFonts runFonts363 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize313 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties75.Append(runFonts363);
            paragraphMarkRunProperties75.Append(fontSize313);

            paragraphProperties75.Append(numberingProperties21);
            paragraphProperties75.Append(snapToGrid67);
            paragraphProperties75.Append(spacingBetweenLines21);
            paragraphProperties75.Append(indentation41);
            paragraphProperties75.Append(paragraphMarkRunProperties75);

            Run run289 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties289 = new RunProperties();
            RunFonts runFonts364 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize314 = new FontSize() { Val = "28" };

            runProperties289.Append(runFonts364);
            runProperties289.Append(fontSize314);
            Text text174 = new Text();
            text174.Text = "Internal Policy & Procedures, External Regulations:";

            run289.Append(runProperties289);
            run289.Append(text174);

            paragraph75.Append(paragraphProperties75);
            paragraph75.Append(run289);

            Paragraph paragraph76 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "004A60EE", ParagraphId = "0F190980", TextId = "77777777" };

            ParagraphProperties paragraphProperties76 = new ParagraphProperties();

            NumberingProperties numberingProperties22 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference22 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId22 = new NumberingId() { Val = 10 };

            numberingProperties22.Append(numberingLevelReference22);
            numberingProperties22.Append(numberingId22);
            SnapToGrid snapToGrid68 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation42 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties76 = new ParagraphMarkRunProperties();
            RunFonts runFonts365 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize315 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties76.Append(runFonts365);
            paragraphMarkRunProperties76.Append(fontSize315);

            paragraphProperties76.Append(numberingProperties22);
            paragraphProperties76.Append(snapToGrid68);
            paragraphProperties76.Append(spacingBetweenLines22);
            paragraphProperties76.Append(indentation42);
            paragraphProperties76.Append(paragraphMarkRunProperties76);

            Run run290 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties290 = new RunProperties();
            RunFonts runFonts366 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize316 = new FontSize() { Val = "28" };

            runProperties290.Append(runFonts366);
            runProperties290.Append(fontSize316);
            Text text175 = new Text();
            text175.Text = "O";

            run290.Append(runProperties290);
            run290.Append(text175);

            Run run291 = new Run() { RsidRunProperties = "0096027D" };

            RunProperties runProperties291 = new RunProperties();
            RunFonts runFonts367 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize317 = new FontSize() { Val = "28" };

            runProperties291.Append(runFonts367);
            runProperties291.Append(fontSize317);
            Text text176 = new Text();
            text176.Text = "thers:";

            run291.Append(runProperties291);
            run291.Append(text176);

            paragraph76.Append(paragraphProperties76);
            paragraph76.Append(run290);
            paragraph76.Append(run291);

            Paragraph paragraph77 = new Paragraph() { RsidParagraphMarkRevision = "0096027D", RsidParagraphAddition = "00200048", RsidRunAdditionDefault = "00200048", ParagraphId = "4E8E3DC1", TextId = "77777777" };

            ParagraphProperties paragraphProperties77 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties77 = new ParagraphMarkRunProperties();
            RunFonts runFonts368 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

            paragraphMarkRunProperties77.Append(runFonts368);

            paragraphProperties77.Append(paragraphMarkRunProperties77);

            paragraph77.Append(paragraphProperties77);

            SectionProperties sectionProperties2 = new SectionProperties() { RsidRPr = "0096027D", RsidR = "00200048", RsidSect = "003B1BF1" };
            PageSize pageSize2 = new PageSize() { Width = (UInt32Value)11906U, Height = (UInt32Value)16838U };
            PageMargin pageMargin2 = new PageMargin() { Top = 1701, Right = (UInt32Value)1077U, Bottom = 1440, Left = (UInt32Value)1077U, Header = (UInt32Value)680U, Footer = (UInt32Value)567U, Gutter = (UInt32Value)0U };
            Columns columns2 = new Columns() { Space = "425" };
            DocGrid docGrid2 = new DocGrid() { Type = DocGridValues.Lines, LinePitch = 360 };

            sectionProperties2.Append(pageSize2);
            sectionProperties2.Append(pageMargin2);
            sectionProperties2.Append(columns2);
            sectionProperties2.Append(docGrid2);

            body1.Append(paragraph1);
            body1.Append(paragraph2);
            body1.Append(table1);
            body1.Append(paragraph17);
            body1.Append(paragraph18);
            body1.Append(table2);
            body1.Append(paragraph31);
            body1.Append(paragraph32);
            body1.Append(paragraph33);
            body1.Append(paragraph34);
            body1.Append(paragraph35);
            body1.Append(paragraph36);
            body1.Append(paragraph37);
            body1.Append(paragraph38);
            body1.Append(paragraph39);
            body1.Append(paragraph40);
            body1.Append(paragraph41);
            body1.Append(paragraph42);
            body1.Append(paragraph43);
            body1.Append(paragraph44);
            body1.Append(paragraph45);
            body1.Append(paragraph46);
            body1.Append(paragraph47);
            body1.Append(paragraph48);
            body1.Append(table3);
            body1.Append(paragraph63);
            body1.Append(paragraph64);
            body1.Append(paragraph65);
            body1.Append(paragraph66);
            body1.Append(paragraph67);
            body1.Append(paragraph68);
            body1.Append(paragraph69);
            body1.Append(paragraph70);
            body1.Append(paragraph71);
            body1.Append(paragraph72);
            body1.Append(paragraph73);
            body1.Append(paragraph74);
            body1.Append(paragraph75);
            body1.Append(paragraph76);
            body1.Append(paragraph77);
            body1.Append(sectionProperties2);

            document1.Append(body1);

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

            Div div1 = new Div() { Id = "624046388" };
            BodyDiv bodyDiv1 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv1 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv1 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv1 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv1 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder1 = new DivBorder();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder1.Append(topBorder2);
            divBorder1.Append(leftBorder2);
            divBorder1.Append(bottomBorder2);
            divBorder1.Append(rightBorder2);

            div1.Append(bodyDiv1);
            div1.Append(leftMarginDiv1);
            div1.Append(rightMarginDiv1);
            div1.Append(topMarginDiv1);
            div1.Append(bottomMarginDiv1);
            div1.Append(divBorder1);

            Div div2 = new Div() { Id = "1898391170" };
            BodyDiv bodyDiv2 = new BodyDiv() { Val = true };
            LeftMarginDiv leftMarginDiv2 = new LeftMarginDiv() { Val = "0" };
            RightMarginDiv rightMarginDiv2 = new RightMarginDiv() { Val = "0" };
            TopMarginDiv topMarginDiv2 = new TopMarginDiv() { Val = "0" };
            BottomMarginDiv bottomMarginDiv2 = new BottomMarginDiv() { Val = "0" };

            DivBorder divBorder2 = new DivBorder();
            TopBorder topBorder3 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder3 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder3 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder3 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            divBorder2.Append(topBorder3);
            divBorder2.Append(leftBorder3);
            divBorder2.Append(bottomBorder3);
            divBorder2.Append(rightBorder3);

            div2.Append(bodyDiv2);
            div2.Append(leftMarginDiv2);
            div2.Append(rightMarginDiv2);
            div2.Append(topMarginDiv2);
            div2.Append(bottomMarginDiv2);
            div2.Append(divBorder2);

            divs1.Append(div1);
            divs1.Append(div2);
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(divs1);
            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
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

            Font font1 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C000785B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Wingdings" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "05000000000000000000" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "02" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Auto };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "00000000", UnicodeSignature1 = "10000000", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "80000000", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "標楷體" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "03000509000000000000" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "88" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Script };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "080E0000", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00100001", CodePageSignature1 = "00000000" };

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

            Font font5 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C200247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font5.Append(panose1Number5);
            font5.Append(fontCharSet5);
            font5.Append(fontFamily5);
            font5.Append(pitch5);
            font5.Append(fontSignature5);

            Font font6 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number6 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet6 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily6 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch6 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature6 = new FontSignature() { UnicodeSignature0 = "E4002EFF", UnicodeSignature1 = "C200247B", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

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

        // Generates content of customXmlPart1.
        private void GenerateCustomXmlPart1Content(CustomXmlPart customXmlPart1)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart1.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><p:properties xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"><documentManagement><lcf76f155ced4ddcb4097134ff3c332f xmlns=\"020e0566-15c6-45f6-927c-a14f1387091e\"><Terms xmlns=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"></Terms></lcf76f155ced4ddcb4097134ff3c332f><TaxCatchAll xmlns=\"cb5579a1-84df-45aa-8c56-deace17c375d\" xsi:nil=\"true\"/></documentManagement></p:properties>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart1.
        private void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
        {
            Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem() { ItemId = "{EA84A370-0142-4EBD-A000-650565A2856C}" };
            dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference1 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties" };
            Ds.SchemaReference schemaReference2 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls" };
            Ds.SchemaReference schemaReference3 = new Ds.SchemaReference() { Uri = "020e0566-15c6-45f6-927c-a14f1387091e" };
            Ds.SchemaReference schemaReference4 = new Ds.SchemaReference() { Uri = "cb5579a1-84df-45aa-8c56-deace17c375d" };

            schemaReferences1.Append(schemaReference1);
            schemaReferences1.Append(schemaReference2);
            schemaReferences1.Append(schemaReference3);
            schemaReferences1.Append(schemaReference4);

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
            Zoom zoom1 = new Zoom() { Percent = "100" };
            BordersDoNotSurroundHeader bordersDoNotSurroundHeader1 = new BordersDoNotSurroundHeader();
            BordersDoNotSurroundFooter bordersDoNotSurroundFooter1 = new BordersDoNotSurroundFooter();
            HideSpellingErrors hideSpellingErrors1 = new HideSpellingErrors();
            HideGrammaticalErrors hideGrammaticalErrors1 = new HideGrammaticalErrors();
            ActiveWritingStyle activeWritingStyle1 = new ActiveWritingStyle() { Language = "en-US", VendorID = (UInt16Value)64U, DllVersion = 6, NaturalLanguageGrammarCheck = true, CheckStyle = false, ApplicationName = "MSWord" };
            ActiveWritingStyle activeWritingStyle2 = new ActiveWritingStyle() { Language = "zh-TW", VendorID = (UInt16Value)64U, DllVersion = 5, NaturalLanguageGrammarCheck = true, CheckStyle = true, ApplicationName = "MSWord" };
            ActiveWritingStyle activeWritingStyle3 = new ActiveWritingStyle() { Language = "zh-HK", VendorID = (UInt16Value)64U, DllVersion = 5, NaturalLanguageGrammarCheck = true, CheckStyle = true, ApplicationName = "MSWord" };
            ActiveWritingStyle activeWritingStyle4 = new ActiveWritingStyle() { Language = "zh-TW", VendorID = (UInt16Value)64U, DllVersion = 0, NaturalLanguageGrammarCheck = true, CheckStyle = true, ApplicationName = "MSWord" };
            ActiveWritingStyle activeWritingStyle5 = new ActiveWritingStyle() { Language = "en-US", VendorID = (UInt16Value)64U, DllVersion = 4096, NaturalLanguageGrammarCheck = true, CheckStyle = false, ApplicationName = "MSWord" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            StylePaneFormatFilter stylePaneFormatFilter1 = new StylePaneFormatFilter() { Val = "3F01", AllStyles = true, CustomStyles = false, LatentStyles = false, StylesInUse = false, HeadingStyles = false, NumberingStyles = false, TableStyles = false, DirectFormattingOnRuns = true, DirectFormattingOnParagraphs = true, DirectFormattingOnNumbering = true, DirectFormattingOnTables = true, ClearFormatting = true, Top3HeadingStyles = true, VisibleStyles = false, AlternateStyleNames = false };
            TrackRevisions trackRevisions1 = new TrackRevisions();
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
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "0014137D" };
            Rsid rsid1 = new Rsid() { Val = "00006DDC" };
            Rsid rsid2 = new Rsid() { Val = "00010E46" };
            Rsid rsid3 = new Rsid() { Val = "00013384" };
            Rsid rsid4 = new Rsid() { Val = "00023ED0" };
            Rsid rsid5 = new Rsid() { Val = "00033A39" };
            Rsid rsid6 = new Rsid() { Val = "00042711" };
            Rsid rsid7 = new Rsid() { Val = "00051944" };
            Rsid rsid8 = new Rsid() { Val = "00060A25" };
            Rsid rsid9 = new Rsid() { Val = "00060A54" };
            Rsid rsid10 = new Rsid() { Val = "00066191" };
            Rsid rsid11 = new Rsid() { Val = "00066F3A" };
            Rsid rsid12 = new Rsid() { Val = "00071B7E" };
            Rsid rsid13 = new Rsid() { Val = "00073DED" };
            Rsid rsid14 = new Rsid() { Val = "00080271" };
            Rsid rsid15 = new Rsid() { Val = "00082F71" };
            Rsid rsid16 = new Rsid() { Val = "000840DF" };
            Rsid rsid17 = new Rsid() { Val = "00096F82" };
            Rsid rsid18 = new Rsid() { Val = "000A741B" };
            Rsid rsid19 = new Rsid() { Val = "000B4F5B" };
            Rsid rsid20 = new Rsid() { Val = "000B6332" };
            Rsid rsid21 = new Rsid() { Val = "000B740F" };
            Rsid rsid22 = new Rsid() { Val = "000C4189" };
            Rsid rsid23 = new Rsid() { Val = "000D4BC6" };
            Rsid rsid24 = new Rsid() { Val = "000E0C5E" };
            Rsid rsid25 = new Rsid() { Val = "000E3880" };
            Rsid rsid26 = new Rsid() { Val = "000F43AE" };
            Rsid rsid27 = new Rsid() { Val = "00100CA1" };
            Rsid rsid28 = new Rsid() { Val = "00101E7E" };
            Rsid rsid29 = new Rsid() { Val = "00112CA3" };
            Rsid rsid30 = new Rsid() { Val = "0013265C" };
            Rsid rsid31 = new Rsid() { Val = "0014137D" };
            Rsid rsid32 = new Rsid() { Val = "00143E14" };
            Rsid rsid33 = new Rsid() { Val = "00153D64" };
            Rsid rsid34 = new Rsid() { Val = "001557E5" };
            Rsid rsid35 = new Rsid() { Val = "00157646" };
            Rsid rsid36 = new Rsid() { Val = "001752F7" };
            Rsid rsid37 = new Rsid() { Val = "00193FEF" };
            Rsid rsid38 = new Rsid() { Val = "001A70EC" };
            Rsid rsid39 = new Rsid() { Val = "001B4FF4" };
            Rsid rsid40 = new Rsid() { Val = "001C3E6D" };
            Rsid rsid41 = new Rsid() { Val = "001C40CD" };
            Rsid rsid42 = new Rsid() { Val = "001C7370" };
            Rsid rsid43 = new Rsid() { Val = "00200048" };
            Rsid rsid44 = new Rsid() { Val = "00210C81" };
            Rsid rsid45 = new Rsid() { Val = "00213B89" };
            Rsid rsid46 = new Rsid() { Val = "00214C36" };
            Rsid rsid47 = new Rsid() { Val = "00215F4D" };
            Rsid rsid48 = new Rsid() { Val = "00216752" };
            Rsid rsid49 = new Rsid() { Val = "00220E02" };
            Rsid rsid50 = new Rsid() { Val = "00222D0C" };
            Rsid rsid51 = new Rsid() { Val = "00231BC8" };
            Rsid rsid52 = new Rsid() { Val = "00241F56" };
            Rsid rsid53 = new Rsid() { Val = "00242440" };
            Rsid rsid54 = new Rsid() { Val = "00243370" };
            Rsid rsid55 = new Rsid() { Val = "0028281B" };
            Rsid rsid56 = new Rsid() { Val = "002856C4" };
            Rsid rsid57 = new Rsid() { Val = "002B63CC" };
            Rsid rsid58 = new Rsid() { Val = "002E4BA9" };
            Rsid rsid59 = new Rsid() { Val = "003174D0" };
            Rsid rsid60 = new Rsid() { Val = "00324D5C" };
            Rsid rsid61 = new Rsid() { Val = "0034608C" };
            Rsid rsid62 = new Rsid() { Val = "00350E12" };
            Rsid rsid63 = new Rsid() { Val = "003619F7" };
            Rsid rsid64 = new Rsid() { Val = "003627CF" };
            Rsid rsid65 = new Rsid() { Val = "00366DF0" };
            Rsid rsid66 = new Rsid() { Val = "00373FDA" };
            Rsid rsid67 = new Rsid() { Val = "00374F61" };
            Rsid rsid68 = new Rsid() { Val = "00375D0B" };
            Rsid rsid69 = new Rsid() { Val = "003B03C7" };
            Rsid rsid70 = new Rsid() { Val = "003B1BF1" };
            Rsid rsid71 = new Rsid() { Val = "003C3877" };
            Rsid rsid72 = new Rsid() { Val = "003E1028" };
            Rsid rsid73 = new Rsid() { Val = "003F2222" };
            Rsid rsid74 = new Rsid() { Val = "003F3517" };
            Rsid rsid75 = new Rsid() { Val = "003F3D6D" };
            Rsid rsid76 = new Rsid() { Val = "003F405A" };
            Rsid rsid77 = new Rsid() { Val = "003F7F8B" };
            Rsid rsid78 = new Rsid() { Val = "0040055A" };
            Rsid rsid79 = new Rsid() { Val = "00405678" };
            Rsid rsid80 = new Rsid() { Val = "00434830" };
            Rsid rsid81 = new Rsid() { Val = "004359B7" };
            Rsid rsid82 = new Rsid() { Val = "004506C8" };
            Rsid rsid83 = new Rsid() { Val = "00485F68" };
            Rsid rsid84 = new Rsid() { Val = "004869E9" };
            Rsid rsid85 = new Rsid() { Val = "004920B3" };
            Rsid rsid86 = new Rsid() { Val = "004924A8" };
            Rsid rsid87 = new Rsid() { Val = "004A60EE" };
            Rsid rsid88 = new Rsid() { Val = "004C2CFC" };
            Rsid rsid89 = new Rsid() { Val = "004C41E7" };
            Rsid rsid90 = new Rsid() { Val = "004C7B6F" };
            Rsid rsid91 = new Rsid() { Val = "004E428D" };
            Rsid rsid92 = new Rsid() { Val = "0052521E" };
            Rsid rsid93 = new Rsid() { Val = "00534826" };
            Rsid rsid94 = new Rsid() { Val = "00567A5E" };
            Rsid rsid95 = new Rsid() { Val = "0057504C" };
            Rsid rsid96 = new Rsid() { Val = "00575F07" };
            Rsid rsid97 = new Rsid() { Val = "00585A36" };
            Rsid rsid98 = new Rsid() { Val = "0059173A" };
            Rsid rsid99 = new Rsid() { Val = "00595527" };
            Rsid rsid100 = new Rsid() { Val = "005C14F0" };
            Rsid rsid101 = new Rsid() { Val = "005C19FA" };
            Rsid rsid102 = new Rsid() { Val = "005D00FB" };
            Rsid rsid103 = new Rsid() { Val = "005D39C8" };
            Rsid rsid104 = new Rsid() { Val = "005E10F2" };
            Rsid rsid105 = new Rsid() { Val = "005E5162" };
            Rsid rsid106 = new Rsid() { Val = "005F213D" };
            Rsid rsid107 = new Rsid() { Val = "006336E9" };
            Rsid rsid108 = new Rsid() { Val = "00646300" };
            Rsid rsid109 = new Rsid() { Val = "00664C80" };
            Rsid rsid110 = new Rsid() { Val = "00673A05" };
            Rsid rsid111 = new Rsid() { Val = "00676306" };
            Rsid rsid112 = new Rsid() { Val = "00686778" };
            Rsid rsid113 = new Rsid() { Val = "006A5363" };
            Rsid rsid114 = new Rsid() { Val = "006C3779" };
            Rsid rsid115 = new Rsid() { Val = "006D2C24" };
            Rsid rsid116 = new Rsid() { Val = "006D564F" };
            Rsid rsid117 = new Rsid() { Val = "006D6273" };
            Rsid rsid118 = new Rsid() { Val = "006F53A7" };
            Rsid rsid119 = new Rsid() { Val = "006F6412" };
            Rsid rsid120 = new Rsid() { Val = "006F6838" };
            Rsid rsid121 = new Rsid() { Val = "00704D62" };
            Rsid rsid122 = new Rsid() { Val = "00710D41" };
            Rsid rsid123 = new Rsid() { Val = "007129F1" };
            Rsid rsid124 = new Rsid() { Val = "00713DAB" };
            Rsid rsid125 = new Rsid() { Val = "00715EA3" };
            Rsid rsid126 = new Rsid() { Val = "00716ABD" };
            Rsid rsid127 = new Rsid() { Val = "00723783" };
            Rsid rsid128 = new Rsid() { Val = "00736443" };
            Rsid rsid129 = new Rsid() { Val = "00746DA3" };
            Rsid rsid130 = new Rsid() { Val = "00760421" };
            Rsid rsid131 = new Rsid() { Val = "007729DD" };
            Rsid rsid132 = new Rsid() { Val = "00787422" };
            Rsid rsid133 = new Rsid() { Val = "0079412A" };
            Rsid rsid134 = new Rsid() { Val = "007C2362" };
            Rsid rsid135 = new Rsid() { Val = "007D1D9B" };
            Rsid rsid136 = new Rsid() { Val = "007E743C" };
            Rsid rsid137 = new Rsid() { Val = "007F22B3" };
            Rsid rsid138 = new Rsid() { Val = "00800493" };
            Rsid rsid139 = new Rsid() { Val = "00801F2D" };
            Rsid rsid140 = new Rsid() { Val = "008410C6" };
            Rsid rsid141 = new Rsid() { Val = "008512BE" };
            Rsid rsid142 = new Rsid() { Val = "008625F1" };
            Rsid rsid143 = new Rsid() { Val = "00866C86" };
            Rsid rsid144 = new Rsid() { Val = "00880D84" };
            Rsid rsid145 = new Rsid() { Val = "00882386" };
            Rsid rsid146 = new Rsid() { Val = "0089312B" };
            Rsid rsid147 = new Rsid() { Val = "00894750" };
            Rsid rsid148 = new Rsid() { Val = "0089528C" };
            Rsid rsid149 = new Rsid() { Val = "00896C9C" };
            Rsid rsid150 = new Rsid() { Val = "008A1C1B" };
            Rsid rsid151 = new Rsid() { Val = "008C261F" };
            Rsid rsid152 = new Rsid() { Val = "008C28B2" };
            Rsid rsid153 = new Rsid() { Val = "008D16E9" };
            Rsid rsid154 = new Rsid() { Val = "008D1E32" };
            Rsid rsid155 = new Rsid() { Val = "009023C8" };
            Rsid rsid156 = new Rsid() { Val = "00904F8B" };
            Rsid rsid157 = new Rsid() { Val = "00917606" };
            Rsid rsid158 = new Rsid() { Val = "00920A29" };
            Rsid rsid159 = new Rsid() { Val = "00932310" };
            Rsid rsid160 = new Rsid() { Val = "009360DF" };
            Rsid rsid161 = new Rsid() { Val = "00936B01" };
            Rsid rsid162 = new Rsid() { Val = "0096027D" };
            Rsid rsid163 = new Rsid() { Val = "0096615A" };
            Rsid rsid164 = new Rsid() { Val = "00972E50" };
            Rsid rsid165 = new Rsid() { Val = "0099123D" };
            Rsid rsid166 = new Rsid() { Val = "0099617D" };
            Rsid rsid167 = new Rsid() { Val = "009A6A85" };
            Rsid rsid168 = new Rsid() { Val = "009A6EFB" };
            Rsid rsid169 = new Rsid() { Val = "009B7102" };
            Rsid rsid170 = new Rsid() { Val = "00A049BB" };
            Rsid rsid171 = new Rsid() { Val = "00A244BE" };
            Rsid rsid172 = new Rsid() { Val = "00A352E0" };
            Rsid rsid173 = new Rsid() { Val = "00A41DB3" };
            Rsid rsid174 = new Rsid() { Val = "00A46B0D" };
            Rsid rsid175 = new Rsid() { Val = "00A566A6" };
            Rsid rsid176 = new Rsid() { Val = "00A663F6" };
            Rsid rsid177 = new Rsid() { Val = "00A72D01" };
            Rsid rsid178 = new Rsid() { Val = "00A94093" };
            Rsid rsid179 = new Rsid() { Val = "00A96EFD" };
            Rsid rsid180 = new Rsid() { Val = "00AA4392" };
            Rsid rsid181 = new Rsid() { Val = "00AD2AF1" };
            Rsid rsid182 = new Rsid() { Val = "00AD3A9D" };
            Rsid rsid183 = new Rsid() { Val = "00AE1DA4" };
            Rsid rsid184 = new Rsid() { Val = "00AF1B88" };
            Rsid rsid185 = new Rsid() { Val = "00B3199E" };
            Rsid rsid186 = new Rsid() { Val = "00B4042C" };
            Rsid rsid187 = new Rsid() { Val = "00B42E31" };
            Rsid rsid188 = new Rsid() { Val = "00B52FB7" };
            Rsid rsid189 = new Rsid() { Val = "00B65A03" };
            Rsid rsid190 = new Rsid() { Val = "00B772CB" };
            Rsid rsid191 = new Rsid() { Val = "00BA581D" };
            Rsid rsid192 = new Rsid() { Val = "00BB5086" };
            Rsid rsid193 = new Rsid() { Val = "00BB7518" };
            Rsid rsid194 = new Rsid() { Val = "00BE42A7" };
            Rsid rsid195 = new Rsid() { Val = "00C245D2" };
            Rsid rsid196 = new Rsid() { Val = "00C34568" };
            Rsid rsid197 = new Rsid() { Val = "00C412A4" };
            Rsid rsid198 = new Rsid() { Val = "00C413B0" };
            Rsid rsid199 = new Rsid() { Val = "00C44C5F" };
            Rsid rsid200 = new Rsid() { Val = "00C56711" };
            Rsid rsid201 = new Rsid() { Val = "00C62CB5" };
            Rsid rsid202 = new Rsid() { Val = "00C671DA" };
            Rsid rsid203 = new Rsid() { Val = "00C7460E" };
            Rsid rsid204 = new Rsid() { Val = "00C74F7A" };
            Rsid rsid205 = new Rsid() { Val = "00C753E4" };
            Rsid rsid206 = new Rsid() { Val = "00C75D1B" };
            Rsid rsid207 = new Rsid() { Val = "00C86D49" };
            Rsid rsid208 = new Rsid() { Val = "00C90DDE" };
            Rsid rsid209 = new Rsid() { Val = "00CA1F96" };
            Rsid rsid210 = new Rsid() { Val = "00CA7B0D" };
            Rsid rsid211 = new Rsid() { Val = "00CB0C7F" };
            Rsid rsid212 = new Rsid() { Val = "00CD2680" };
            Rsid rsid213 = new Rsid() { Val = "00CE480A" };
            Rsid rsid214 = new Rsid() { Val = "00CF2F8F" };
            Rsid rsid215 = new Rsid() { Val = "00CF43E8" };
            Rsid rsid216 = new Rsid() { Val = "00CF64BF" };
            Rsid rsid217 = new Rsid() { Val = "00D02A33" };
            Rsid rsid218 = new Rsid() { Val = "00D36ED8" };
            Rsid rsid219 = new Rsid() { Val = "00D41898" };
            Rsid rsid220 = new Rsid() { Val = "00D43514" };
            Rsid rsid221 = new Rsid() { Val = "00D62F73" };
            Rsid rsid222 = new Rsid() { Val = "00D66E75" };
            Rsid rsid223 = new Rsid() { Val = "00D73633" };
            Rsid rsid224 = new Rsid() { Val = "00D7709C" };
            Rsid rsid225 = new Rsid() { Val = "00D9623A" };
            Rsid rsid226 = new Rsid() { Val = "00DB24E5" };
            Rsid rsid227 = new Rsid() { Val = "00DD3CDC" };
            Rsid rsid228 = new Rsid() { Val = "00DD4C3C" };
            Rsid rsid229 = new Rsid() { Val = "00DE479D" };
            Rsid rsid230 = new Rsid() { Val = "00DE4990" };
            Rsid rsid231 = new Rsid() { Val = "00E2524D" };
            Rsid rsid232 = new Rsid() { Val = "00E30AF0" };
            Rsid rsid233 = new Rsid() { Val = "00E35208" };
            Rsid rsid234 = new Rsid() { Val = "00E519A7" };
            Rsid rsid235 = new Rsid() { Val = "00E55A94" };
            Rsid rsid236 = new Rsid() { Val = "00E659D9" };
            Rsid rsid237 = new Rsid() { Val = "00E6762F" };
            Rsid rsid238 = new Rsid() { Val = "00EB5D9F" };
            Rsid rsid239 = new Rsid() { Val = "00ED5245" };
            Rsid rsid240 = new Rsid() { Val = "00ED6DFF" };
            Rsid rsid241 = new Rsid() { Val = "00ED7D88" };
            Rsid rsid242 = new Rsid() { Val = "00EE6462" };
            Rsid rsid243 = new Rsid() { Val = "00F03C81" };
            Rsid rsid244 = new Rsid() { Val = "00F4377B" };
            Rsid rsid245 = new Rsid() { Val = "00F45650" };
            Rsid rsid246 = new Rsid() { Val = "00F45CFE" };
            Rsid rsid247 = new Rsid() { Val = "00F469D5" };
            Rsid rsid248 = new Rsid() { Val = "00F55A18" };
            Rsid rsid249 = new Rsid() { Val = "00F6052A" };
            Rsid rsid250 = new Rsid() { Val = "00F64F71" };
            Rsid rsid251 = new Rsid() { Val = "00F7611F" };
            Rsid rsid252 = new Rsid() { Val = "00F87EAC" };
            Rsid rsid253 = new Rsid() { Val = "00FA0623" };
            Rsid rsid254 = new Rsid() { Val = "00FC58E2" };
            Rsid rsid255 = new Rsid() { Val = "00FC6F9F" };
            Rsid rsid256 = new Rsid() { Val = "00FD24E6" };
            Rsid rsid257 = new Rsid() { Val = "00FD4180" };
            Rsid rsid258 = new Rsid() { Val = "00FD740C" };
            Rsid rsid259 = new Rsid() { Val = "00FE2DFB" };
            Rsid rsid260 = new Rsid() { Val = "00FE4552" };
            Rsid rsid261 = new Rsid() { Val = "00FF6CCA" };

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
            rsids1.Append(rsid114);
            rsids1.Append(rsid115);
            rsids1.Append(rsid116);
            rsids1.Append(rsid117);
            rsids1.Append(rsid118);
            rsids1.Append(rsid119);
            rsids1.Append(rsid120);
            rsids1.Append(rsid121);
            rsids1.Append(rsid122);
            rsids1.Append(rsid123);
            rsids1.Append(rsid124);
            rsids1.Append(rsid125);
            rsids1.Append(rsid126);
            rsids1.Append(rsid127);
            rsids1.Append(rsid128);
            rsids1.Append(rsid129);
            rsids1.Append(rsid130);
            rsids1.Append(rsid131);
            rsids1.Append(rsid132);
            rsids1.Append(rsid133);
            rsids1.Append(rsid134);
            rsids1.Append(rsid135);
            rsids1.Append(rsid136);
            rsids1.Append(rsid137);
            rsids1.Append(rsid138);
            rsids1.Append(rsid139);
            rsids1.Append(rsid140);
            rsids1.Append(rsid141);
            rsids1.Append(rsid142);
            rsids1.Append(rsid143);
            rsids1.Append(rsid144);
            rsids1.Append(rsid145);
            rsids1.Append(rsid146);
            rsids1.Append(rsid147);
            rsids1.Append(rsid148);
            rsids1.Append(rsid149);
            rsids1.Append(rsid150);
            rsids1.Append(rsid151);
            rsids1.Append(rsid152);
            rsids1.Append(rsid153);
            rsids1.Append(rsid154);
            rsids1.Append(rsid155);
            rsids1.Append(rsid156);
            rsids1.Append(rsid157);
            rsids1.Append(rsid158);
            rsids1.Append(rsid159);
            rsids1.Append(rsid160);
            rsids1.Append(rsid161);
            rsids1.Append(rsid162);
            rsids1.Append(rsid163);
            rsids1.Append(rsid164);
            rsids1.Append(rsid165);
            rsids1.Append(rsid166);
            rsids1.Append(rsid167);
            rsids1.Append(rsid168);
            rsids1.Append(rsid169);
            rsids1.Append(rsid170);
            rsids1.Append(rsid171);
            rsids1.Append(rsid172);
            rsids1.Append(rsid173);
            rsids1.Append(rsid174);
            rsids1.Append(rsid175);
            rsids1.Append(rsid176);
            rsids1.Append(rsid177);
            rsids1.Append(rsid178);
            rsids1.Append(rsid179);
            rsids1.Append(rsid180);
            rsids1.Append(rsid181);
            rsids1.Append(rsid182);
            rsids1.Append(rsid183);
            rsids1.Append(rsid184);
            rsids1.Append(rsid185);
            rsids1.Append(rsid186);
            rsids1.Append(rsid187);
            rsids1.Append(rsid188);
            rsids1.Append(rsid189);
            rsids1.Append(rsid190);
            rsids1.Append(rsid191);
            rsids1.Append(rsid192);
            rsids1.Append(rsid193);
            rsids1.Append(rsid194);
            rsids1.Append(rsid195);
            rsids1.Append(rsid196);
            rsids1.Append(rsid197);
            rsids1.Append(rsid198);
            rsids1.Append(rsid199);
            rsids1.Append(rsid200);
            rsids1.Append(rsid201);
            rsids1.Append(rsid202);
            rsids1.Append(rsid203);
            rsids1.Append(rsid204);
            rsids1.Append(rsid205);
            rsids1.Append(rsid206);
            rsids1.Append(rsid207);
            rsids1.Append(rsid208);
            rsids1.Append(rsid209);
            rsids1.Append(rsid210);
            rsids1.Append(rsid211);
            rsids1.Append(rsid212);
            rsids1.Append(rsid213);
            rsids1.Append(rsid214);
            rsids1.Append(rsid215);
            rsids1.Append(rsid216);
            rsids1.Append(rsid217);
            rsids1.Append(rsid218);
            rsids1.Append(rsid219);
            rsids1.Append(rsid220);
            rsids1.Append(rsid221);
            rsids1.Append(rsid222);
            rsids1.Append(rsid223);
            rsids1.Append(rsid224);
            rsids1.Append(rsid225);
            rsids1.Append(rsid226);
            rsids1.Append(rsid227);
            rsids1.Append(rsid228);
            rsids1.Append(rsid229);
            rsids1.Append(rsid230);
            rsids1.Append(rsid231);
            rsids1.Append(rsid232);
            rsids1.Append(rsid233);
            rsids1.Append(rsid234);
            rsids1.Append(rsid235);
            rsids1.Append(rsid236);
            rsids1.Append(rsid237);
            rsids1.Append(rsid238);
            rsids1.Append(rsid239);
            rsids1.Append(rsid240);
            rsids1.Append(rsid241);
            rsids1.Append(rsid242);
            rsids1.Append(rsid243);
            rsids1.Append(rsid244);
            rsids1.Append(rsid245);
            rsids1.Append(rsid246);
            rsids1.Append(rsid247);
            rsids1.Append(rsid248);
            rsids1.Append(rsid249);
            rsids1.Append(rsid250);
            rsids1.Append(rsid251);
            rsids1.Append(rsid252);
            rsids1.Append(rsid253);
            rsids1.Append(rsid254);
            rsids1.Append(rsid255);
            rsids1.Append(rsid256);
            rsids1.Append(rsid257);
            rsids1.Append(rsid258);
            rsids1.Append(rsid259);
            rsids1.Append(rsid260);
            rsids1.Append(rsid261);

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
            DoNotIncludeSubdocsInStats doNotIncludeSubdocsInStats1 = new DoNotIncludeSubdocsInStats();

            ShapeDefaults shapeDefaults2 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2050 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "2" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults2.Append(shapeDefaults3);
            shapeDefaults2.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "18A27263" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{1B93F51E-8762-4C6A-83CD-7EB07B4CD619}" };

            settings1.Append(zoom1);
            settings1.Append(bordersDoNotSurroundHeader1);
            settings1.Append(bordersDoNotSurroundFooter1);
            settings1.Append(hideSpellingErrors1);
            settings1.Append(hideGrammaticalErrors1);
            settings1.Append(activeWritingStyle1);
            settings1.Append(activeWritingStyle2);
            settings1.Append(activeWritingStyle3);
            settings1.Append(activeWritingStyle4);
            settings1.Append(activeWritingStyle5);
            settings1.Append(proofState1);
            settings1.Append(stylePaneFormatFilter1);
            settings1.Append(trackRevisions1);
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
            settings1.Append(doNotIncludeSubdocsInStats1);
            settings1.Append(shapeDefaults2);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);
            settings1.Append(documentId1);
            settings1.Append(chartTrackingRefBased1);
            settings1.Append(persistentDocumentId1);

            documentSettingsPart1.Settings = settings1;
        }

        // Generates content of footerPart1.
        private void GenerateFooterPart1Content(FooterPart footerPart1)
        {
            Footer footer1 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            footer1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            footer1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            footer1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            footer1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            footer1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            footer1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            footer1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            footer1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            footer1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            footer1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            footer1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            footer1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            footer1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            footer1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            footer1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            footer1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            footer1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph78 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "000E3880", RsidParagraphProperties = "000E3880", RsidRunAdditionDefault = "000E3880", ParagraphId = "3CDD7DAB", TextId = "34F7D39B" };

            ParagraphProperties paragraphProperties78 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "a4" };

            Tabs tabs9 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Clear, Position = 8306 };
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Right, Position = 9752 };

            tabs9.Append(tabStop9);
            tabs9.Append(tabStop10);

            paragraphProperties78.Append(paragraphStyleId7);
            paragraphProperties78.Append(tabs9);

            DeletedRun deletedRun65 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "53" };

            Run run292 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties292 = new RunProperties();
            Color color168 = new Color() { Val = "0000FF" };

            runProperties292.Append(color168);
            DeletedText deletedText113 = new DeletedText();
            deletedText113.Text = "[";

            run292.Append(runProperties292);
            run292.Append(deletedText113);

            Run run293 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties293 = new RunProperties();
            RunFonts runFonts369 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color169 = new Color() { Val = "0000FF" };

            runProperties293.Append(runFonts369);
            runProperties293.Append(color169);
            DeletedText deletedText114 = new DeletedText();
            deletedText114.Text = "組織";

            run293.Append(runProperties293);
            run293.Append(deletedText114);

            Run run294 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties294 = new RunProperties();
            Color color170 = new Color() { Val = "0000FF" };

            runProperties294.Append(color170);
            DeletedText deletedText115 = new DeletedText();
            deletedText115.Text = "].[";

            run294.Append(runProperties294);
            run294.Append(deletedText115);

            deletedRun65.Append(run292);
            deletedRun65.Append(run293);
            deletedRun65.Append(run294);

            Run run295 = new Run();

            RunProperties runProperties295 = new RunProperties();
            RunFonts runFonts370 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color171 = new Color() { Val = "0000FF" };

            runProperties295.Append(runFonts370);
            runProperties295.Append(color171);
            Text text177 = new Text();
            text177.Text = dt.Rows[0]["CompanyName"].ToString();

            run295.Append(runProperties295);
            run295.Append(text177);

            DeletedRun deletedRun66 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "54" };

            Run run296 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties296 = new RunProperties();
            Color color172 = new Color() { Val = "0000FF" };

            runProperties296.Append(color172);
            DeletedText deletedText116 = new DeletedText();
            deletedText116.Text = "]";

            run296.Append(runProperties296);
            run296.Append(deletedText116);

            deletedRun66.Append(run296);

            Run run297 = new Run();

            RunProperties runProperties297 = new RunProperties();
            Color color173 = new Color() { Val = "0000FF" };

            runProperties297.Append(color173);
            Text text178 = new Text();
            text178.Text = "_";

            run297.Append(runProperties297);
            run297.Append(text178);

            DeletedRun deletedRun67 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "55" };

            Run run298 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties298 = new RunProperties();
            Color color174 = new Color() { Val = "0000FF" };

            runProperties298.Append(color174);
            DeletedText deletedText117 = new DeletedText();
            deletedText117.Text = "[";

            run298.Append(runProperties298);
            run298.Append(deletedText117);

            Run run299 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties299 = new RunProperties();
            RunFonts runFonts371 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color175 = new Color() { Val = "0000FF" };

            runProperties299.Append(runFonts371);
            runProperties299.Append(color175);
            DeletedText deletedText118 = new DeletedText();
            deletedText118.Text = "查程";

            run299.Append(runProperties299);
            run299.Append(deletedText118);

            Run run300 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties300 = new RunProperties();
            Color color176 = new Color() { Val = "0000FF" };

            runProperties300.Append(color176);
            DeletedText deletedText119 = new DeletedText();
            deletedText119.Text = "].[";

            run300.Append(runProperties300);
            run300.Append(deletedText119);

            deletedRun67.Append(run298);
            deletedRun67.Append(run299);
            deletedRun67.Append(run300);

            Run run301 = new Run();

            RunProperties runProperties301 = new RunProperties();
            RunFonts runFonts372 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color177 = new Color() { Val = "0000FF" };

            runProperties301.Append(runFonts372);
            runProperties301.Append(color177);
            Text text179 = new Text();
            text179.Text = "查程編號";

            run301.Append(runProperties301);
            run301.Append(text179);

            DeletedRun deletedRun68 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "56" };

            Run run302 = new Run() { RsidRunDeletion = "001C7370" };

            RunProperties runProperties302 = new RunProperties();
            Color color178 = new Color() { Val = "0000FF" };

            runProperties302.Append(color178);
            DeletedText deletedText120 = new DeletedText();
            deletedText120.Text = "]";

            run302.Append(runProperties302);
            run302.Append(deletedText120);

            deletedRun68.Append(run302);

            Run run303 = new Run();

            RunProperties runProperties303 = new RunProperties();
            Color color179 = new Color() { Val = "0000FF" };

            runProperties303.Append(color179);
            TabChar tabChar1 = new TabChar();

            run303.Append(runProperties303);
            run303.Append(tabChar1);

            Run run304 = new Run();

            RunProperties runProperties304 = new RunProperties();
            Color color180 = new Color() { Val = "0000FF" };

            runProperties304.Append(color180);
            TabChar tabChar2 = new TabChar();

            run304.Append(runProperties304);
            run304.Append(tabChar2);

            Run run305 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run305.Append(fieldChar1);

            Run run306 = new Run();
            FieldCode fieldCode1 = new FieldCode();
            fieldCode1.Text = "PAGE   \\* MERGEFORMAT";

            run306.Append(fieldCode1);

            Run run307 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run307.Append(fieldChar2);

            Run run308 = new Run() { RsidRunProperties = "00B772CB", RsidRunAddition = "00B772CB" };

            RunProperties runProperties305 = new RunProperties();
            NoProof noProof1 = new NoProof();
            Languages languages33 = new Languages() { Val = "zh-TW" };

            runProperties305.Append(noProof1);
            runProperties305.Append(languages33);
            Text text180 = new Text();
            text180.Text = "2";

            run308.Append(runProperties305);
            run308.Append(text180);

            Run run309 = new Run();
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run309.Append(fieldChar3);

            paragraph78.Append(paragraphProperties78);
            paragraph78.Append(deletedRun65);
            paragraph78.Append(run295);
            paragraph78.Append(deletedRun66);
            paragraph78.Append(run297);
            paragraph78.Append(deletedRun67);
            paragraph78.Append(run301);
            paragraph78.Append(deletedRun68);
            paragraph78.Append(run303);
            paragraph78.Append(run304);
            paragraph78.Append(run305);
            paragraph78.Append(run306);
            paragraph78.Append(run307);
            paragraph78.Append(run308);
            paragraph78.Append(run309);

            footer1.Append(paragraph78);

            footerPart1.Footer = footer1;
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
            Ds.DataStoreItem dataStoreItem2 = new Ds.DataStoreItem() { ItemId = "{F0B9C9E9-37C5-4183-9690-4F5ED3708956}" };
            dataStoreItem2.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences2 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference5 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/sharepoint/v3/contenttype/forms" };

            schemaReferences2.Append(schemaReference5);

            dataStoreItem2.Append(schemaReferences2);

            customXmlPropertiesPart2.DataStoreItem = dataStoreItem2;
        }

        // Generates content of customXmlPart3.
        private void GenerateCustomXmlPart3Content(CustomXmlPart customXmlPart3)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart3.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><ct:contentTypeSchema ct:_=\"\" ma:_=\"\" ma:contentTypeName=\"文件\" ma:contentTypeID=\"0x010100D465D9393F8B7B429C8C9CE9DCEB2AA6\" ma:contentTypeVersion=\"9\" ma:contentTypeDescription=\"建立新的文件。\" ma:contentTypeScope=\"\" ma:versionID=\"1e5f73f4dc337242fde86b74da489c82\" xmlns:ct=\"http://schemas.microsoft.com/office/2006/metadata/contentType\" xmlns:ma=\"http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes\">\r\n<xsd:schema targetNamespace=\"http://schemas.microsoft.com/office/2006/metadata/properties\" ma:root=\"true\" ma:fieldsID=\"8c42c8921edb5ff244e8d475413dc0a6\" ns2:_=\"\" ns3:_=\"\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" xmlns:ns2=\"020e0566-15c6-45f6-927c-a14f1387091e\" xmlns:ns3=\"cb5579a1-84df-45aa-8c56-deace17c375d\">\r\n<xsd:import namespace=\"020e0566-15c6-45f6-927c-a14f1387091e\"/>\r\n<xsd:import namespace=\"cb5579a1-84df-45aa-8c56-deace17c375d\"/>\r\n<xsd:element name=\"properties\">\r\n<xsd:complexType>\r\n<xsd:sequence>\r\n<xsd:element name=\"documentManagement\">\r\n<xsd:complexType>\r\n<xsd:all>\r\n<xsd:element ref=\"ns2:MediaServiceMetadata\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceFastMetadata\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceObjectDetectorVersions\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:lcf76f155ced4ddcb4097134ff3c332f\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns3:TaxCatchAll\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceOCR\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceGenerationTime\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceEventHashCode\" minOccurs=\"0\"/>\r\n</xsd:all>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:sequence>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"020e0566-15c6-45f6-927c-a14f1387091e\" elementFormDefault=\"qualified\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:dms=\"http://schemas.microsoft.com/office/2006/documentManagement/types\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\">\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/2006/documentManagement/types\"/>\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"/>\r\n<xsd:element name=\"MediaServiceMetadata\" ma:index=\"8\" nillable=\"true\" ma:displayName=\"MediaServiceMetadata\" ma:hidden=\"true\" ma:internalName=\"MediaServiceMetadata\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceFastMetadata\" ma:index=\"9\" nillable=\"true\" ma:displayName=\"MediaServiceFastMetadata\" ma:hidden=\"true\" ma:internalName=\"MediaServiceFastMetadata\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceObjectDetectorVersions\" ma:index=\"10\" nillable=\"true\" ma:displayName=\"MediaServiceObjectDetectorVersions\" ma:hidden=\"true\" ma:indexed=\"true\" ma:internalName=\"MediaServiceObjectDetectorVersions\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"lcf76f155ced4ddcb4097134ff3c332f\" ma:index=\"12\" nillable=\"true\" ma:taxonomy=\"true\" ma:internalName=\"lcf76f155ced4ddcb4097134ff3c332f\" ma:taxonomyFieldName=\"MediaServiceImageTags\" ma:displayName=\"影像標籤\" ma:readOnly=\"false\" ma:fieldId=\"{5cf76f15-5ced-4ddc-b409-7134ff3c332f}\" ma:taxonomyMulti=\"true\" ma:sspId=\"9a4d633b-e48a-43c2-b595-a1cd44314b02\" ma:termSetId=\"09814cd3-568e-fe90-9814-8d621ff8fb84\" ma:anchorId=\"fba54fb3-c3e1-fe81-a776-ca4b69148c4d\" ma:open=\"true\" ma:isKeyword=\"false\">\r\n<xsd:complexType>\r\n<xsd:sequence>\r\n<xsd:element ref=\"pc:Terms\" minOccurs=\"0\" maxOccurs=\"1\"></xsd:element>\r\n</xsd:sequence>\r\n</xsd:complexType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceOCR\" ma:index=\"14\" nillable=\"true\" ma:displayName=\"Extracted Text\" ma:internalName=\"MediaServiceOCR\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\">\r\n<xsd:maxLength value=\"255\"/>\r\n</xsd:restriction>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceGenerationTime\" ma:index=\"15\" nillable=\"true\" ma:displayName=\"MediaServiceGenerationTime\" ma:hidden=\"true\" ma:internalName=\"MediaServiceGenerationTime\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceEventHashCode\" ma:index=\"16\" nillable=\"true\" ma:displayName=\"MediaServiceEventHashCode\" ma:hidden=\"true\" ma:internalName=\"MediaServiceEventHashCode\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"cb5579a1-84df-45aa-8c56-deace17c375d\" elementFormDefault=\"qualified\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:dms=\"http://schemas.microsoft.com/office/2006/documentManagement/types\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\">\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/2006/documentManagement/types\"/>\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"/>\r\n<xsd:element name=\"TaxCatchAll\" ma:index=\"13\" nillable=\"true\" ma:displayName=\"Taxonomy Catch All Column\" ma:hidden=\"true\" ma:list=\"{c0d4e7c3-a6fa-49eb-b1fa-7b786d1ce50c}\" ma:internalName=\"TaxCatchAll\" ma:showField=\"CatchAllData\" ma:web=\"cb5579a1-84df-45aa-8c56-deace17c375d\">\r\n<xsd:complexType>\r\n<xsd:complexContent>\r\n<xsd:extension base=\"dms:MultiChoiceLookup\">\r\n<xsd:sequence>\r\n<xsd:element name=\"Value\" type=\"dms:Lookup\" maxOccurs=\"unbounded\" minOccurs=\"0\" nillable=\"true\"/>\r\n</xsd:sequence>\r\n</xsd:extension>\r\n</xsd:complexContent>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" elementFormDefault=\"qualified\" attributeFormDefault=\"unqualified\" blockDefault=\"#all\" xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:odoc=\"http://schemas.microsoft.com/internal/obd\">\r\n<xsd:import namespace=\"http://purl.org/dc/elements/1.1/\" schemaLocation=\"http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd\"/>\r\n<xsd:import namespace=\"http://purl.org/dc/terms/\" schemaLocation=\"http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcterms.xsd\"/>\r\n<xsd:element name=\"coreProperties\" type=\"CT_coreProperties\"/>\r\n<xsd:complexType name=\"CT_coreProperties\">\r\n<xsd:all>\r\n<xsd:element ref=\"dc:creator\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dcterms:created\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dc:identifier\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"contentType\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\" ma:index=\"0\" ma:displayName=\"內容類型\"/>\r\n<xsd:element ref=\"dc:title\" minOccurs=\"0\" maxOccurs=\"1\" ma:index=\"4\" ma:displayName=\"標題\"/>\r\n<xsd:element ref=\"dc:subject\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dc:description\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"keywords\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element ref=\"dc:language\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"category\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element name=\"version\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element name=\"revision\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\">\r\n<xsd:annotation>\r\n<xsd:documentation>\r\n                        This value indicates the number of saves or revisions. The application is responsible for updating this value after each revision.\r\n                    </xsd:documentation>\r\n</xsd:annotation>\r\n</xsd:element>\r\n<xsd:element name=\"lastModifiedBy\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element ref=\"dcterms:modified\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"contentStatus\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n</xsd:all>\r\n</xsd:complexType>\r\n</xsd:schema>\r\n<xs:schema targetNamespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\" elementFormDefault=\"qualified\" attributeFormDefault=\"unqualified\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">\r\n<xs:element name=\"Person\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:DisplayName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:AccountId\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:AccountType\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"DisplayName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"AccountId\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"AccountType\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"BDCAssociatedEntity\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:BDCEntity\" minOccurs=\"0\" maxOccurs=\"unbounded\"></xs:element>\r\n</xs:sequence>\r\n<xs:attribute ref=\"pc:EntityNamespace\"></xs:attribute>\r\n<xs:attribute ref=\"pc:EntityName\"></xs:attribute>\r\n<xs:attribute ref=\"pc:SystemInstanceName\"></xs:attribute>\r\n<xs:attribute ref=\"pc:AssociationName\"></xs:attribute>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:attribute name=\"EntityNamespace\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"EntityName\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"SystemInstanceName\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"AssociationName\" type=\"xs:string\"></xs:attribute>\r\n<xs:element name=\"BDCEntity\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:EntityDisplayName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityInstanceReference\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId1\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId2\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId3\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId4\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId5\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"EntityDisplayName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityInstanceReference\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId1\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId2\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId3\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId4\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId5\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"Terms\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:TermInfo\" minOccurs=\"0\" maxOccurs=\"unbounded\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"TermInfo\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:TermName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:TermId\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"TermName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"TermId\" type=\"xs:string\"></xs:element>\r\n</xs:schema>\r\n</ct:contentTypeSchema>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart3.
        private void GenerateCustomXmlPropertiesPart3Content(CustomXmlPropertiesPart customXmlPropertiesPart3)
        {
            Ds.DataStoreItem dataStoreItem3 = new Ds.DataStoreItem() { ItemId = "{09834855-5DB9-4FA4-A817-F70411810EDF}" };
            dataStoreItem3.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences3 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference6 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/contentType" };
            Ds.SchemaReference schemaReference7 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes" };
            Ds.SchemaReference schemaReference8 = new Ds.SchemaReference() { Uri = "http://www.w3.org/2001/XMLSchema" };
            Ds.SchemaReference schemaReference9 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties" };
            Ds.SchemaReference schemaReference10 = new Ds.SchemaReference() { Uri = "020e0566-15c6-45f6-927c-a14f1387091e" };
            Ds.SchemaReference schemaReference11 = new Ds.SchemaReference() { Uri = "cb5579a1-84df-45aa-8c56-deace17c375d" };
            Ds.SchemaReference schemaReference12 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/documentManagement/types" };
            Ds.SchemaReference schemaReference13 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls" };
            Ds.SchemaReference schemaReference14 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" };
            Ds.SchemaReference schemaReference15 = new Ds.SchemaReference() { Uri = "http://purl.org/dc/elements/1.1/" };
            Ds.SchemaReference schemaReference16 = new Ds.SchemaReference() { Uri = "http://purl.org/dc/terms/" };
            Ds.SchemaReference schemaReference17 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/internal/obd" };

            schemaReferences3.Append(schemaReference6);
            schemaReferences3.Append(schemaReference7);
            schemaReferences3.Append(schemaReference8);
            schemaReferences3.Append(schemaReference9);
            schemaReferences3.Append(schemaReference10);
            schemaReferences3.Append(schemaReference11);
            schemaReferences3.Append(schemaReference12);
            schemaReferences3.Append(schemaReference13);
            schemaReferences3.Append(schemaReference14);
            schemaReferences3.Append(schemaReference15);
            schemaReferences3.Append(schemaReference16);
            schemaReferences3.Append(schemaReference17);

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
            RunFonts runFonts373 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "新細明體", ComplexScript = "Times New Roman" };
            Languages languages34 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts373);
            runPropertiesBaseStyle1.Append(languages34);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);
            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 0, DefaultSemiHidden = false, DefaultUnhideWhenUsed = false, DefaultPrimaryStyle = false, Count = 376 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "footnote text", UiPriority = 99 };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "footer", UiPriority = 99 };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "caption", SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "Title", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "Subtitle", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "Strong", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "Emphasis", PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "Normal Table", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "Table Simple 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "Table Simple 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Table Simple 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Table Classic 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Table Classic 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Table Classic 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Table Classic 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Table Colorful 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Table Colorful 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "Table Colorful 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Table Columns 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Table Columns 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Table Columns 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Table Columns 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Table Columns 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Table Grid 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Table Grid 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Table Grid 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Table Grid 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Table Grid 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "Table Grid 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Table Grid 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Table Grid 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Table List 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Table List 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Table List 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Table List 4", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Table List 5", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Table List 6", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Table List 7", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Table List 8", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Table 3D effects 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Table Contemporary", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Table Elegant", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Table Professional", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Table Subtle 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Table Subtle 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Table Web 1", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Table Web 2", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Table Web 3", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Table Theme", SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Revision", UiPriority = 99, SemiHidden = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo138 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo139 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo140 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo141 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo142 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo143 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo144 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo145 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo146 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo147 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo148 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo149 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo150 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo151 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo152 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60 };
            LatentStyleExceptionInfo latentStyleExceptionInfo153 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61 };
            LatentStyleExceptionInfo latentStyleExceptionInfo154 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62 };
            LatentStyleExceptionInfo latentStyleExceptionInfo155 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63 };
            LatentStyleExceptionInfo latentStyleExceptionInfo156 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64 };
            LatentStyleExceptionInfo latentStyleExceptionInfo157 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65 };
            LatentStyleExceptionInfo latentStyleExceptionInfo158 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66 };
            LatentStyleExceptionInfo latentStyleExceptionInfo159 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67 };
            LatentStyleExceptionInfo latentStyleExceptionInfo160 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68 };
            LatentStyleExceptionInfo latentStyleExceptionInfo161 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69 };
            LatentStyleExceptionInfo latentStyleExceptionInfo162 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70 };
            LatentStyleExceptionInfo latentStyleExceptionInfo163 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71 };
            LatentStyleExceptionInfo latentStyleExceptionInfo164 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72 };
            LatentStyleExceptionInfo latentStyleExceptionInfo165 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73 };
            LatentStyleExceptionInfo latentStyleExceptionInfo166 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo167 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo168 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo169 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo170 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo171 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo172 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, SemiHidden = true, UnhideWhenUsed = true, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo173 = new LatentStyleExceptionInfo() { Name = "Plain Table 1", UiPriority = 41 };
            LatentStyleExceptionInfo latentStyleExceptionInfo174 = new LatentStyleExceptionInfo() { Name = "Plain Table 2", UiPriority = 42 };
            LatentStyleExceptionInfo latentStyleExceptionInfo175 = new LatentStyleExceptionInfo() { Name = "Plain Table 3", UiPriority = 43 };
            LatentStyleExceptionInfo latentStyleExceptionInfo176 = new LatentStyleExceptionInfo() { Name = "Plain Table 4", UiPriority = 44 };
            LatentStyleExceptionInfo latentStyleExceptionInfo177 = new LatentStyleExceptionInfo() { Name = "Plain Table 5", UiPriority = 45 };
            LatentStyleExceptionInfo latentStyleExceptionInfo178 = new LatentStyleExceptionInfo() { Name = "Grid Table Light", UiPriority = 40 };
            LatentStyleExceptionInfo latentStyleExceptionInfo179 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo180 = new LatentStyleExceptionInfo() { Name = "Grid Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo181 = new LatentStyleExceptionInfo() { Name = "Grid Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo182 = new LatentStyleExceptionInfo() { Name = "Grid Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo183 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo184 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo185 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo186 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo187 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo188 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo189 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo190 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo191 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo192 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo193 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo194 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo195 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo196 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo197 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo198 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo199 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo200 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo201 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo202 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo203 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo204 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo205 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo206 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo207 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo208 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo209 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo210 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo211 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo212 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo213 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo214 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo215 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo216 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo217 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo218 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo219 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo220 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo221 = new LatentStyleExceptionInfo() { Name = "Grid Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo222 = new LatentStyleExceptionInfo() { Name = "Grid Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo223 = new LatentStyleExceptionInfo() { Name = "Grid Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo224 = new LatentStyleExceptionInfo() { Name = "Grid Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo225 = new LatentStyleExceptionInfo() { Name = "Grid Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo226 = new LatentStyleExceptionInfo() { Name = "Grid Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo227 = new LatentStyleExceptionInfo() { Name = "Grid Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo228 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo229 = new LatentStyleExceptionInfo() { Name = "List Table 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo230 = new LatentStyleExceptionInfo() { Name = "List Table 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo231 = new LatentStyleExceptionInfo() { Name = "List Table 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo232 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo233 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo234 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo235 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 1", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo236 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 1", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo237 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 1", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo238 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 1", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo239 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 1", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo240 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 1", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo241 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 1", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo242 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 2", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo243 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 2", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo244 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 2", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo245 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 2", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo246 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 2", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo247 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 2", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo248 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 2", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo249 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 3", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo250 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 3", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo251 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 3", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo252 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 3", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo253 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 3", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo254 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 3", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo255 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 3", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo256 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 4", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo257 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 4", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo258 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 4", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo259 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 4", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo260 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 4", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo261 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 4", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo262 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 4", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo263 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 5", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo264 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 5", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo265 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 5", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo266 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 5", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo267 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 5", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo268 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 5", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo269 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 5", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo270 = new LatentStyleExceptionInfo() { Name = "List Table 1 Light Accent 6", UiPriority = 46 };
            LatentStyleExceptionInfo latentStyleExceptionInfo271 = new LatentStyleExceptionInfo() { Name = "List Table 2 Accent 6", UiPriority = 47 };
            LatentStyleExceptionInfo latentStyleExceptionInfo272 = new LatentStyleExceptionInfo() { Name = "List Table 3 Accent 6", UiPriority = 48 };
            LatentStyleExceptionInfo latentStyleExceptionInfo273 = new LatentStyleExceptionInfo() { Name = "List Table 4 Accent 6", UiPriority = 49 };
            LatentStyleExceptionInfo latentStyleExceptionInfo274 = new LatentStyleExceptionInfo() { Name = "List Table 5 Dark Accent 6", UiPriority = 50 };
            LatentStyleExceptionInfo latentStyleExceptionInfo275 = new LatentStyleExceptionInfo() { Name = "List Table 6 Colorful Accent 6", UiPriority = 51 };
            LatentStyleExceptionInfo latentStyleExceptionInfo276 = new LatentStyleExceptionInfo() { Name = "List Table 7 Colorful Accent 6", UiPriority = 52 };
            LatentStyleExceptionInfo latentStyleExceptionInfo277 = new LatentStyleExceptionInfo() { Name = "Mention", UiPriority = 99, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo278 = new LatentStyleExceptionInfo() { Name = "Smart Hyperlink", UiPriority = 99, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo279 = new LatentStyleExceptionInfo() { Name = "Hashtag", UiPriority = 99, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo280 = new LatentStyleExceptionInfo() { Name = "Unresolved Mention", UiPriority = 99, SemiHidden = true, UnhideWhenUsed = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo281 = new LatentStyleExceptionInfo() { Name = "Smart Link", UiPriority = 99, SemiHidden = true, UnhideWhenUsed = true };

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

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "a", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid262 = new Rsid() { Val = "0014137D" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl() { Val = false };

            styleParagraphProperties1.Append(widowControl1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts374 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern1 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize318 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(runFonts374);
            styleRunProperties1.Append(kern1);
            styleRunProperties1.Append(fontSize318);
            styleRunProperties1.Append(fontSizeComplexScript54);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid262);
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

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation4 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault4 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin4 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin4 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault4.Append(topMargin1);
            tableCellMarginDefault4.Append(tableCellLeftMargin4);
            tableCellMarginDefault4.Append(bottomMargin1);
            tableCellMarginDefault4.Append(tableCellRightMargin4);

            styleTableProperties1.Append(tableIndentation4);
            styleTableProperties1.Append(tableCellMarginDefault4);

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
            StyleName styleName5 = new StyleName() { Val = "header" };
            BasedOn basedOn1 = new BasedOn() { Val = "a" };
            Rsid rsid263 = new Rsid() { Val = "0014137D" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs10 = new Tabs();
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs10.Append(tabStop11);
            tabs10.Append(tabStop12);
            SnapToGrid snapToGrid69 = new SnapToGrid() { Val = false };

            styleParagraphProperties2.Append(tabs10);
            styleParagraphProperties2.Append(snapToGrid69);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            FontSize fontSize319 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties2.Append(fontSize319);
            styleRunProperties2.Append(fontSizeComplexScript55);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(rsid263);
            style5.Append(styleParagraphProperties2);
            style5.Append(styleRunProperties2);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "a4" };
            StyleName styleName6 = new StyleName() { Val = "footer" };
            BasedOn basedOn2 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "a5" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            Rsid rsid264 = new Rsid() { Val = "003C3877" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs11 = new Tabs();
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs11.Append(tabStop13);
            tabs11.Append(tabStop14);
            SnapToGrid snapToGrid70 = new SnapToGrid() { Val = false };

            styleParagraphProperties3.Append(tabs11);
            styleParagraphProperties3.Append(snapToGrid70);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            FontSize fontSize320 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties3.Append(fontSize320);
            styleRunProperties3.Append(fontSizeComplexScript56);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle1);
            style6.Append(uIPriority4);
            style6.Append(rsid264);
            style6.Append(styleParagraphProperties3);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "a5", CustomStyle = true };
            StyleName styleName7 = new StyleName() { Val = "頁尾 字元" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a4" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            Rsid rsid265 = new Rsid() { Val = "003C3877" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts375 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern2 = new Kern() { Val = (UInt32Value)2U };

            styleRunProperties4.Append(runFonts375);
            styleRunProperties4.Append(kern2);

            style7.Append(styleName7);
            style7.Append(linkedStyle2);
            style7.Append(uIPriority5);
            style7.Append(rsid265);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() { Type = StyleValues.Table, StyleId = "a6" };
            StyleName styleName8 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn3 = new BasedOn() { Val = "a1" };
            Rsid rsid266 = new Rsid() { Val = "00882386" };

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder4 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder4 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder4 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder4 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder4);
            tableBorders2.Append(leftBorder4);
            tableBorders2.Append(bottomBorder4);
            tableBorders2.Append(rightBorder4);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);

            styleTableProperties2.Append(tableBorders2);

            style8.Append(styleName8);
            style8.Append(basedOn3);
            style8.Append(rsid266);
            style8.Append(styleTableProperties2);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "a7" };
            StyleName styleName9 = new StyleName() { Val = "footnote text" };
            BasedOn basedOn4 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "a8" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            Rsid rsid267 = new Rsid() { Val = "006C3779" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            SnapToGrid snapToGrid71 = new SnapToGrid() { Val = false };

            styleParagraphProperties4.Append(snapToGrid71);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            FontSize fontSize321 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties5.Append(fontSize321);
            styleRunProperties5.Append(fontSizeComplexScript57);

            style9.Append(styleName9);
            style9.Append(basedOn4);
            style9.Append(linkedStyle3);
            style9.Append(uIPriority6);
            style9.Append(rsid267);
            style9.Append(styleParagraphProperties4);
            style9.Append(styleRunProperties5);

            Style style10 = new Style() { Type = StyleValues.Character, StyleId = "a8", CustomStyle = true };
            StyleName styleName10 = new StyleName() { Val = "註腳文字 字元" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a7" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            Rsid rsid268 = new Rsid() { Val = "006C3779" };

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts376 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern3 = new Kern() { Val = (UInt32Value)2U };

            styleRunProperties6.Append(runFonts376);
            styleRunProperties6.Append(kern3);

            style10.Append(styleName10);
            style10.Append(linkedStyle4);
            style10.Append(uIPriority7);
            style10.Append(rsid268);
            style10.Append(styleRunProperties6);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "a9" };
            StyleName styleName11 = new StyleName() { Val = "footnote reference" };
            Rsid rsid269 = new Rsid() { Val = "006C3779" };

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            VerticalTextAlignment verticalTextAlignment3 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            styleRunProperties7.Append(verticalTextAlignment3);

            style11.Append(styleName11);
            style11.Append(rsid269);
            style11.Append(styleRunProperties7);

            Style style12 = new Style() { Type = StyleValues.Character, StyleId = "aa" };
            StyleName styleName12 = new StyleName() { Val = "annotation reference" };
            Rsid rsid270 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            FontSize fontSize322 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties8.Append(fontSize322);
            styleRunProperties8.Append(fontSizeComplexScript58);

            style12.Append(styleName12);
            style12.Append(rsid270);
            style12.Append(styleRunProperties8);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "ab" };
            StyleName styleName13 = new StyleName() { Val = "annotation text" };
            BasedOn basedOn5 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "ac" };
            Rsid rsid271 = new Rsid() { Val = "00214C36" };

            style13.Append(styleName13);
            style13.Append(basedOn5);
            style13.Append(linkedStyle5);
            style13.Append(rsid271);

            Style style14 = new Style() { Type = StyleValues.Character, StyleId = "ac", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "註解文字 字元" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "ab" };
            Rsid rsid272 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts377 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern4 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize323 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties9.Append(runFonts377);
            styleRunProperties9.Append(kern4);
            styleRunProperties9.Append(fontSize323);
            styleRunProperties9.Append(fontSizeComplexScript59);

            style14.Append(styleName14);
            style14.Append(linkedStyle6);
            style14.Append(rsid272);
            style14.Append(styleRunProperties9);

            Style style15 = new Style() { Type = StyleValues.Paragraph, StyleId = "ad" };
            StyleName styleName15 = new StyleName() { Val = "annotation subject" };
            BasedOn basedOn6 = new BasedOn() { Val = "ab" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "ab" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "ae" };
            Rsid rsid273 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            Bold bold282 = new Bold();
            BoldComplexScript boldComplexScript145 = new BoldComplexScript();

            styleRunProperties10.Append(bold282);
            styleRunProperties10.Append(boldComplexScript145);

            style15.Append(styleName15);
            style15.Append(basedOn6);
            style15.Append(nextParagraphStyle1);
            style15.Append(linkedStyle7);
            style15.Append(rsid273);
            style15.Append(styleRunProperties10);

            Style style16 = new Style() { Type = StyleValues.Character, StyleId = "ae", CustomStyle = true };
            StyleName styleName16 = new StyleName() { Val = "註解主旨 字元" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "ad" };
            Rsid rsid274 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts378 = new RunFonts() { EastAsia = "標楷體" };
            Bold bold283 = new Bold();
            BoldComplexScript boldComplexScript146 = new BoldComplexScript();
            Kern kern5 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize324 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties11.Append(runFonts378);
            styleRunProperties11.Append(bold283);
            styleRunProperties11.Append(boldComplexScript146);
            styleRunProperties11.Append(kern5);
            styleRunProperties11.Append(fontSize324);
            styleRunProperties11.Append(fontSizeComplexScript60);

            style16.Append(styleName16);
            style16.Append(linkedStyle8);
            style16.Append(rsid274);
            style16.Append(styleRunProperties11);

            Style style17 = new Style() { Type = StyleValues.Paragraph, StyleId = "af" };
            StyleName styleName17 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn7 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "af0" };
            Rsid rsid275 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts379 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "新細明體" };
            FontSize fontSize325 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties12.Append(runFonts379);
            styleRunProperties12.Append(fontSize325);
            styleRunProperties12.Append(fontSizeComplexScript61);

            style17.Append(styleName17);
            style17.Append(basedOn7);
            style17.Append(linkedStyle9);
            style17.Append(rsid275);
            style17.Append(styleRunProperties12);

            Style style18 = new Style() { Type = StyleValues.Character, StyleId = "af0", CustomStyle = true };
            StyleName styleName18 = new StyleName() { Val = "註解方塊文字 字元" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "af" };
            Rsid rsid276 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts380 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "新細明體", ComplexScript = "Times New Roman" };
            Kern kern6 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize326 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties13.Append(runFonts380);
            styleRunProperties13.Append(kern6);
            styleRunProperties13.Append(fontSize326);
            styleRunProperties13.Append(fontSizeComplexScript62);

            style18.Append(styleName18);
            style18.Append(linkedStyle10);
            style18.Append(rsid276);
            style18.Append(styleRunProperties13);

            Style style19 = new Style() { Type = StyleValues.Character, StyleId = "normaltextrun", CustomStyle = true };
            StyleName styleName19 = new StyleName() { Val = "normaltextrun" };
            Rsid rsid277 = new Rsid() { Val = "00EB5D9F" };

            style19.Append(styleName19);
            style19.Append(rsid277);

            Style style20 = new Style() { Type = StyleValues.Paragraph, StyleId = "af1" };
            StyleName styleName20 = new StyleName() { Val = "Revision" };
            StyleHidden styleHidden1 = new StyleHidden();
            UIPriority uIPriority8 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            Rsid rsid278 = new Rsid() { Val = "001C7370" };

            StyleRunProperties styleRunProperties14 = new StyleRunProperties();
            RunFonts runFonts381 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern7 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize327 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties14.Append(runFonts381);
            styleRunProperties14.Append(kern7);
            styleRunProperties14.Append(fontSize327);
            styleRunProperties14.Append(fontSizeComplexScript63);

            style20.Append(styleName20);
            style20.Append(styleHidden1);
            style20.Append(uIPriority8);
            style20.Append(semiHidden4);
            style20.Append(rsid278);
            style20.Append(styleRunProperties14);

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
            styles1.Append(style20);

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

            Paragraph paragraph79 = new Paragraph() { RsidParagraphMarkRevision = "00213B89", RsidParagraphAddition = "00882386", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00220E02", ParagraphId = "654419DC", TextId = "7D723CDC" };

            ParagraphProperties paragraphProperties79 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties79.Append(paragraphStyleId8);

            Run run310 = new Run();

            RunProperties runProperties306 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties306.Append(noProof2);

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251657728U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "682C1017", AnchorId = "2BFFBE01" };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "3086100";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "-100965";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 3159760L, Cy = 356235L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)9U, Name = "文字方塊 9" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks();

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 3159760L, Cy = 356235L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline() { Width = 6350 };
            A.NoFill noFill2 = new A.NoFill();

            outline1.Append(noFill2);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(outline1);

            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph80 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "3C5ED4B3", TextId = "77777777" };

            ParagraphProperties paragraphProperties80 = new ParagraphProperties();
            SnapToGrid snapToGrid72 = new SnapToGrid() { Val = false };
            Justification justification45 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties78 = new ParagraphMarkRunProperties();
            RunFonts runFonts382 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize328 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "20" };
            Languages languages35 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties78.Append(runFonts382);
            paragraphMarkRunProperties78.Append(fontSize328);
            paragraphMarkRunProperties78.Append(fontSizeComplexScript64);
            paragraphMarkRunProperties78.Append(languages35);

            paragraphProperties80.Append(snapToGrid72);
            paragraphProperties80.Append(justification45);
            paragraphProperties80.Append(paragraphMarkRunProperties78);

            Run run311 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties307 = new RunProperties();
            RunFonts runFonts383 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize329 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "20" };

            runProperties307.Append(runFonts383);
            runProperties307.Append(fontSize329);
            runProperties307.Append(fontSizeComplexScript65);
            Text text181 = new Text();
            text181.Text = "Information Classification: Confidential";

            run311.Append(runProperties307);
            run311.Append(text181);

            paragraph80.Append(paragraphProperties80);
            paragraph80.Append(run311);

            Paragraph paragraph81 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "1D12BD9B", TextId = "22E95DC1" };

            ParagraphProperties paragraphProperties81 = new ParagraphProperties();
            SnapToGrid snapToGrid73 = new SnapToGrid() { Val = false };
            Justification justification46 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties79 = new ParagraphMarkRunProperties();
            RunFonts runFonts384 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize330 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties79.Append(runFonts384);
            paragraphMarkRunProperties79.Append(fontSize330);
            paragraphMarkRunProperties79.Append(fontSizeComplexScript66);

            paragraphProperties81.Append(snapToGrid73);
            paragraphProperties81.Append(justification46);
            paragraphProperties81.Append(paragraphMarkRunProperties79);

            Run run312 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties308 = new RunProperties();
            RunFonts runFonts385 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize331 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "20" };

            runProperties308.Append(runFonts385);
            runProperties308.Append(fontSize331);
            runProperties308.Append(fontSizeComplexScript67);
            Text text182 = new Text();
            text182.Text = "Declassified Date:";

            run312.Append(runProperties308);
            run312.Append(text182);

            Run run313 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties309 = new RunProperties();
            RunFonts runFonts386 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color181 = new Color() { Val = "3333FF" };
            FontSize fontSize332 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "20" };

            runProperties309.Append(runFonts386);
            runProperties309.Append(color181);
            runProperties309.Append(fontSize332);
            runProperties309.Append(fontSizeComplexScript68);
            Text text183 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text183.Text = " ";

            run313.Append(runProperties309);
            run313.Append(text183);

            DeletedRun deletedRun69 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "45" };

            Run run314 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties310 = new RunProperties();
            RunFonts runFonts387 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color182 = new Color() { Val = "3333FF" };
            FontSize fontSize333 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "20" };

            runProperties310.Append(runFonts387);
            runProperties310.Append(color182);
            runProperties310.Append(fontSize333);
            runProperties310.Append(fontSizeComplexScript69);
            DeletedText deletedText121 = new DeletedText();
            deletedText121.Text = "[";

            run314.Append(runProperties310);
            run314.Append(deletedText121);

            Run run315 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties311 = new RunProperties();
            RunFonts runFonts388 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color183 = new Color() { Val = "3333FF" };
            FontSize fontSize334 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "20" };
            Languages languages36 = new Languages() { EastAsia = "zh-HK" };

            runProperties311.Append(runFonts388);
            runProperties311.Append(color183);
            runProperties311.Append(fontSize334);
            runProperties311.Append(fontSizeComplexScript70);
            runProperties311.Append(languages36);
            DeletedText deletedText122 = new DeletedText();
            deletedText122.Text = "查程";

            run315.Append(runProperties311);
            run315.Append(deletedText122);

            Run run316 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties312 = new RunProperties();
            RunFonts runFonts389 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color184 = new Color() { Val = "3333FF" };
            FontSize fontSize335 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "20" };

            runProperties312.Append(runFonts389);
            runProperties312.Append(color184);
            runProperties312.Append(fontSize335);
            runProperties312.Append(fontSizeComplexScript71);
            DeletedText deletedText123 = new DeletedText();
            deletedText123.Text = "table].[";

            run316.Append(runProperties312);
            run316.Append(deletedText123);

            deletedRun69.Append(run314);
            deletedRun69.Append(run315);
            deletedRun69.Append(run316);

            Run run317 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties313 = new RunProperties();
            RunFonts runFonts390 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color185 = new Color() { Val = "3333FF" };
            FontSize fontSize336 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "20" };
            Languages languages37 = new Languages() { EastAsia = "zh-HK" };

            runProperties313.Append(runFonts390);
            runProperties313.Append(color185);
            runProperties313.Append(fontSize336);
            runProperties313.Append(fontSizeComplexScript72);
            runProperties313.Append(languages37);
            Text text184 = new Text();
            //text184.Text = "查核迄日";
            if (DateTime.TryParse(dt.Rows[0]["Ptitleenddate"].ToString(), out DateTime date5))
            {
                text184.Text = date5.ToString("yyyy-MM-dd");
            }
            else
            {
                text184.Text = "";
            }

            run317.Append(runProperties313);
            run317.Append(text184);

            InsertedRun insertedRun1 = new InsertedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T16:10:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "46" };

            Run run318 = new Run() { RsidRunAddition = "00080271" };

            RunProperties runProperties314 = new RunProperties();
            RunFonts runFonts391 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color186 = new Color() { Val = "3333FF" };
            FontSize fontSize337 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "20" };
            Languages languages38 = new Languages() { EastAsia = "zh-HK" };

            runProperties314.Append(runFonts391);
            runProperties314.Append(color186);
            runProperties314.Append(fontSize337);
            runProperties314.Append(fontSizeComplexScript73);
            runProperties314.Append(languages38);
            Text text185 = new Text();
            text185.Text = "解密";

            run318.Append(runProperties314);
            run318.Append(text185);

            insertedRun1.Append(run318);

            DeletedRun deletedRun70 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "47" };

            Run run319 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties315 = new RunProperties();
            RunFonts runFonts392 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color187 = new Color() { Val = "3333FF" };
            FontSize fontSize338 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "20" };

            runProperties315.Append(runFonts392);
            runProperties315.Append(color187);
            runProperties315.Append(fontSize338);
            runProperties315.Append(fontSizeComplexScript74);
            DeletedText deletedText124 = new DeletedText();
            deletedText124.Text = "]";

            run319.Append(runProperties315);
            run319.Append(deletedText124);

            deletedRun70.Append(run319);

            DeletedRun deletedRun71 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T16:10:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "48" };

            Run run320 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "00080271" };

            RunProperties runProperties316 = new RunProperties();
            RunFonts runFonts393 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color188 = new Color() { Val = "3333FF" };
            FontSize fontSize339 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "20" };

            runProperties316.Append(runFonts393);
            runProperties316.Append(color188);
            runProperties316.Append(fontSize339);
            runProperties316.Append(fontSizeComplexScript75);
            DeletedText deletedText125 = new DeletedText();
            deletedText125.Text = "+5";

            run320.Append(runProperties316);
            run320.Append(deletedText125);

            Run run321 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "00080271" };

            RunProperties runProperties317 = new RunProperties();
            RunFonts runFonts394 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color189 = new Color() { Val = "3333FF" };
            FontSize fontSize340 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "20" };

            runProperties317.Append(runFonts394);
            runProperties317.Append(color189);
            runProperties317.Append(fontSize340);
            runProperties317.Append(fontSizeComplexScript76);
            DeletedText deletedText126 = new DeletedText();
            deletedText126.Text = "年";

            run321.Append(runProperties317);
            run321.Append(deletedText126);

            deletedRun71.Append(run320);
            deletedRun71.Append(run321);

            paragraph81.Append(paragraphProperties81);
            paragraph81.Append(run312);
            paragraph81.Append(run313);
            paragraph81.Append(deletedRun69);
            paragraph81.Append(run317);
            paragraph81.Append(insertedRun1);
            paragraph81.Append(deletedRun70);
            paragraph81.Append(deletedRun71);

            Paragraph paragraph82 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "7300B055", TextId = "77777777" };

            ParagraphProperties paragraphProperties82 = new ParagraphProperties();
            SnapToGrid snapToGrid74 = new SnapToGrid() { Val = false };
            Justification justification47 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties82.Append(snapToGrid74);
            paragraphProperties82.Append(justification47);

            paragraph82.Append(paragraphProperties82);

            textBoxContent1.Append(paragraph80);
            textBoxContent1.Append(paragraph81);
            textBoxContent1.Append(paragraph82);

            textBoxInfo21.Append(textBoxContent1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = false, VerticalOverflow = A.TextVerticalOverflowValues.Overflow, HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 0, RightToLeftColumns = false, FromWordArt = false, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, ForceAntiAlias = false, CompatibleLineSpacing = true };

            A.PresetTextWrap presetTextWrap1 = new A.PresetTextWrap() { Preset = A.TextShapeValues.TextNoShape };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetTextWrap1.Append(adjustValueList2);
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(presetTextWrap1);
            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            alternateContentChoice1.Append(drawing1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
            shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "2BFFBE01"));
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            V.Shape shape1 = new V.Shape() { Id = "文字方塊 9", Style = "position:absolute;margin-left:243pt;margin-top:-7.95pt;width:248.8pt;height:28.05pt;z-index:251657728;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top", OptionalString = "_x0000_s1026", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQAzVqvWGAIAADUEAAAOAAAAZHJzL2Uyb0RvYy54bWysU11v0zAUfUfiP1h+p+mHWljUdCqbipCq\nbVKH9uw6dhvh+Jprt8n49Vw7SYsGT4gX58b3+5zj5W1bG3ZW6CuwBZ+MxpwpK6Gs7KHg3543Hz5x\n5oOwpTBgVcFflee3q/fvlo3L1RSOYEqFjIpYnzeu4McQXJ5lXh5VLfwInLLk1IC1CPSLh6xE0VD1\n2mTT8XiRNYClQ5DKe7q975x8leprrWR41NqrwEzBabaQTkznPp7ZainyAwp3rGQ/hviHKWpRWWp6\nKXUvgmAnrP4oVVcSwYMOIwl1BlpXUqUdaJvJ+M02u6NwKu1C4Hh3gcn/v7Ly4bxzT8hC+xlaIjAt\n4d0W5HdP2GSN83kfEzH1uafouGirsY5fWoFRImH7esFTtYFJupxN5jcfF+SS5JvNF9PZPAKeXbMd\n+vBFQc2iUXAkvtIE4rz1oQsdQmIzC5vKmMSZsawp+GI2H6eEi4eKG9sP3s0apw7tvqW0aO6hfKWF\nEToteCc3FTXfCh+eBBL5NC8JOjzSoQ1QE+gtzo6AP/92H+OJE/Jy1pCYCu5/nAQqzsxXS2xF5Q0G\nDsZ+MOypvgPS54SeipPJpAQMZjA1Qv1COl/HLuQSVlKvgofBvAudpOmdSLVepyDSlxNha3dODrxG\nKJ/bF4GuxzsQUw8wyEzkb2DvYjvg16cAukqcXFHscSZtJlb7dxTF//t/irq+9tUvAAAA//8DAFBL\nAwQUAAYACAAAACEAgOgf6eAAAAAKAQAADwAAAGRycy9kb3ducmV2LnhtbEyPzU7DMBCE70i8g7VI\n3Fo7BaI0xKkQPzcotAUJbk5skojYjuxNGt6e5QTH0Yxmvik2s+3ZZELsvJOQLAUw42qvO9dIeD08\nLDJgEZXTqvfOSPg2ETbl6Umhcu2PbmemPTaMSlzMlYQWccg5j3VrrIpLPxhH3qcPViHJ0HAd1JHK\nbc9XQqTcqs7RQqsGc9ua+ms/Wgn9ewyPlcCP6a55wpdnPr7dJ1spz8/mm2tgaGb8C8MvPqFDSUyV\nH52OrJdwmaX0BSUskqs1MEqss4sUWEWWWAEvC/7/QvkDAAD//wMAUEsBAi0AFAAGAAgAAAAhALaD\nOJL+AAAA4QEAABMAAAAAAAAAAAAAAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYA\nCAAAACEAOP0h/9YAAACUAQAACwAAAAAAAAAAAAAAAAAvAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYA\nCAAAACEAM1ar1hgCAAA1BAAADgAAAAAAAAAAAAAAAAAuAgAAZHJzL2Uyb0RvYy54bWxQSwECLQAU\nAAYACAAAACEAgOgf6eAAAAAKAQAADwAAAAAAAAAAAAAAAAByBAAAZHJzL2Rvd25yZXYueG1sUEsF\nBgAAAAAEAAQA8wAAAH8FAAAAAA==\n" };

            V.TextBox textBox1 = new V.TextBox() { Inset = "0,0,0,0" };

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph83 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "3C5ED4B3", TextId = "77777777" };

            ParagraphProperties paragraphProperties83 = new ParagraphProperties();
            SnapToGrid snapToGrid75 = new SnapToGrid() { Val = false };
            Justification justification48 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties80 = new ParagraphMarkRunProperties();
            RunFonts runFonts395 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize341 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "20" };
            Languages languages39 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties80.Append(runFonts395);
            paragraphMarkRunProperties80.Append(fontSize341);
            paragraphMarkRunProperties80.Append(fontSizeComplexScript77);
            paragraphMarkRunProperties80.Append(languages39);

            paragraphProperties83.Append(snapToGrid75);
            paragraphProperties83.Append(justification48);
            paragraphProperties83.Append(paragraphMarkRunProperties80);

            Run run322 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties318 = new RunProperties();
            RunFonts runFonts396 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize342 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "20" };

            runProperties318.Append(runFonts396);
            runProperties318.Append(fontSize342);
            runProperties318.Append(fontSizeComplexScript78);
            Text text186 = new Text();
            text186.Text = "Information Classification: Confidential";

            run322.Append(runProperties318);
            run322.Append(text186);

            paragraph83.Append(paragraphProperties83);
            paragraph83.Append(run322);

            Paragraph paragraph84 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00E659D9", RsidParagraphProperties = "00E659D9", RsidRunAdditionDefault = "00E659D9", ParagraphId = "1D12BD9B", TextId = "22E95DC1" };

            ParagraphProperties paragraphProperties84 = new ParagraphProperties();
            SnapToGrid snapToGrid76 = new SnapToGrid() { Val = false };
            Justification justification49 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties81 = new ParagraphMarkRunProperties();
            RunFonts runFonts397 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize343 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties81.Append(runFonts397);
            paragraphMarkRunProperties81.Append(fontSize343);
            paragraphMarkRunProperties81.Append(fontSizeComplexScript79);

            paragraphProperties84.Append(snapToGrid76);
            paragraphProperties84.Append(justification49);
            paragraphProperties84.Append(paragraphMarkRunProperties81);

            Run run323 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties319 = new RunProperties();
            RunFonts runFonts398 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            FontSize fontSize344 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "20" };

            runProperties319.Append(runFonts398);
            runProperties319.Append(fontSize344);
            runProperties319.Append(fontSizeComplexScript80);
            Text text187 = new Text();
            text187.Text = "Declassified Date:";

            run323.Append(runProperties319);
            run323.Append(text187);

            Run run324 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties320 = new RunProperties();
            RunFonts runFonts399 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color190 = new Color() { Val = "3333FF" };
            FontSize fontSize345 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "20" };

            runProperties320.Append(runFonts399);
            runProperties320.Append(color190);
            runProperties320.Append(fontSize345);
            runProperties320.Append(fontSizeComplexScript81);
            Text text188 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text188.Text = " ";

            run324.Append(runProperties320);
            run324.Append(text188);

            DeletedRun deletedRun72 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "49" };

            Run run325 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties321 = new RunProperties();
            RunFonts runFonts400 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color191 = new Color() { Val = "3333FF" };
            FontSize fontSize346 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "20" };

            runProperties321.Append(runFonts400);
            runProperties321.Append(color191);
            runProperties321.Append(fontSize346);
            runProperties321.Append(fontSizeComplexScript82);
            DeletedText deletedText127 = new DeletedText();
            deletedText127.Text = "[";

            run325.Append(runProperties321);
            run325.Append(deletedText127);

            Run run326 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties322 = new RunProperties();
            RunFonts runFonts401 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color192 = new Color() { Val = "3333FF" };
            FontSize fontSize347 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "20" };
            Languages languages40 = new Languages() { EastAsia = "zh-HK" };

            runProperties322.Append(runFonts401);
            runProperties322.Append(color192);
            runProperties322.Append(fontSize347);
            runProperties322.Append(fontSizeComplexScript83);
            runProperties322.Append(languages40);
            DeletedText deletedText128 = new DeletedText();
            deletedText128.Text = "查程";

            run326.Append(runProperties322);
            run326.Append(deletedText128);

            Run run327 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties323 = new RunProperties();
            RunFonts runFonts402 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color193 = new Color() { Val = "3333FF" };
            FontSize fontSize348 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "20" };

            runProperties323.Append(runFonts402);
            runProperties323.Append(color193);
            runProperties323.Append(fontSize348);
            runProperties323.Append(fontSizeComplexScript84);
            DeletedText deletedText129 = new DeletedText();
            deletedText129.Text = "table].[";

            run327.Append(runProperties323);
            run327.Append(deletedText129);

            deletedRun72.Append(run325);
            deletedRun72.Append(run326);
            deletedRun72.Append(run327);

            Run run328 = new Run() { RsidRunProperties = "00E659D9" };

            RunProperties runProperties324 = new RunProperties();
            RunFonts runFonts403 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color194 = new Color() { Val = "3333FF" };
            FontSize fontSize349 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "20" };
            Languages languages41 = new Languages() { EastAsia = "zh-HK" };

            runProperties324.Append(runFonts403);
            runProperties324.Append(color194);
            runProperties324.Append(fontSize349);
            runProperties324.Append(fontSizeComplexScript85);
            runProperties324.Append(languages41);
            Text text189 = new Text();
            text189.Text = "查核迄日";

            run328.Append(runProperties324);
            run328.Append(text189);

            InsertedRun insertedRun2 = new InsertedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T16:10:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "50" };

            Run run329 = new Run() { RsidRunAddition = "00080271" };

            RunProperties runProperties325 = new RunProperties();
            RunFonts runFonts404 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color195 = new Color() { Val = "3333FF" };
            FontSize fontSize350 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "20" };
            Languages languages42 = new Languages() { EastAsia = "zh-HK" };

            runProperties325.Append(runFonts404);
            runProperties325.Append(color195);
            runProperties325.Append(fontSize350);
            runProperties325.Append(fontSizeComplexScript86);
            runProperties325.Append(languages42);
            Text text190 = new Text();
            text190.Text = "解密";

            run329.Append(runProperties325);
            run329.Append(text190);

            insertedRun2.Append(run329);

            DeletedRun deletedRun73 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T14:04:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "51" };

            Run run330 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "001C7370" };

            RunProperties runProperties326 = new RunProperties();
            RunFonts runFonts405 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color196 = new Color() { Val = "3333FF" };
            FontSize fontSize351 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "20" };

            runProperties326.Append(runFonts405);
            runProperties326.Append(color196);
            runProperties326.Append(fontSize351);
            runProperties326.Append(fontSizeComplexScript87);
            DeletedText deletedText130 = new DeletedText();
            deletedText130.Text = "]";

            run330.Append(runProperties326);
            run330.Append(deletedText130);

            deletedRun73.Append(run330);

            DeletedRun deletedRun74 = new DeletedRun() { Author = "余亭妍", Date = System.Xml.XmlConvert.ToDateTime("2023-08-11T16:10:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind), Id = "52" };

            Run run331 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "00080271" };

            RunProperties runProperties327 = new RunProperties();
            RunFonts runFonts406 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color197 = new Color() { Val = "3333FF" };
            FontSize fontSize352 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "20" };

            runProperties327.Append(runFonts406);
            runProperties327.Append(color197);
            runProperties327.Append(fontSize352);
            runProperties327.Append(fontSizeComplexScript88);
            DeletedText deletedText131 = new DeletedText();
            deletedText131.Text = "+5";

            run331.Append(runProperties327);
            run331.Append(deletedText131);

            Run run332 = new Run() { RsidRunProperties = "00E659D9", RsidRunDeletion = "00080271" };

            RunProperties runProperties328 = new RunProperties();
            RunFonts runFonts407 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Color color198 = new Color() { Val = "3333FF" };
            FontSize fontSize353 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "20" };

            runProperties328.Append(runFonts407);
            runProperties328.Append(color198);
            runProperties328.Append(fontSize353);
            runProperties328.Append(fontSizeComplexScript89);
            DeletedText deletedText132 = new DeletedText();
            deletedText132.Text = "年";

            run332.Append(runProperties328);
            run332.Append(deletedText132);

            deletedRun74.Append(run331);
            deletedRun74.Append(run332);

            paragraph84.Append(paragraphProperties84);
            paragraph84.Append(run323);
            paragraph84.Append(run324);
            paragraph84.Append(deletedRun72);
            paragraph84.Append(run328);
            paragraph84.Append(insertedRun2);
            paragraph84.Append(deletedRun73);
            paragraph84.Append(deletedRun74);

            Paragraph paragraph85 = new Paragraph() { RsidParagraphMarkRevision = "00E659D9", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "7300B055", TextId = "77777777" };

            ParagraphProperties paragraphProperties85 = new ParagraphProperties();
            SnapToGrid snapToGrid77 = new SnapToGrid() { Val = false };
            Justification justification50 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties85.Append(snapToGrid77);
            paragraphProperties85.Append(justification50);

            paragraph85.Append(paragraphProperties85);

            textBoxContent2.Append(paragraph83);
            textBoxContent2.Append(paragraph84);
            textBoxContent2.Append(paragraph85);

            textBox1.Append(textBoxContent2);

            shape1.Append(textBox1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            alternateContentFallback1.Append(picture1);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            run310.Append(runProperties306);
            run310.Append(alternateContent1);

            paragraph79.Append(paragraphProperties79);
            paragraph79.Append(run310);

            header1.Append(paragraph79);

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
            Nsid nsid1 = new Nsid() { Val = "04A90420" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "92F2F87C" };

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "0409001B" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText1 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();

            Tabs tabs12 = new Tabs();
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Number, Position = 1680 };

            tabs12.Append(tabStop15);
            Indentation indentation43 = new Indentation() { Start = "1680", Hanging = "720" };

            previousParagraphProperties1.Append(tabs12);
            previousParagraphProperties1.Append(indentation43);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts408 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties1.Append(runFonts408);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);
            level1.Append(numberingSymbolRunProperties1);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText2 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();

            Tabs tabs13 = new Tabs();
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Number, Position = 1920 };

            tabs13.Append(tabStop16);
            Indentation indentation44 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties2.Append(tabs13);
            previousParagraphProperties2.Append(indentation44);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();

            Tabs tabs14 = new Tabs();
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Number, Position = 2400 };

            tabs14.Append(tabStop17);
            Indentation indentation45 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties3.Append(tabs14);
            previousParagraphProperties3.Append(indentation45);

            level3.Append(startNumberingValue3);
            level3.Append(numberingFormat3);
            level3.Append(levelText3);
            level3.Append(levelJustification3);
            level3.Append(previousParagraphProperties3);

            Level level4 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue4 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat4 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText4 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification4 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties4 = new PreviousParagraphProperties();

            Tabs tabs15 = new Tabs();
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs15.Append(tabStop18);
            Indentation indentation46 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties4.Append(tabs15);
            previousParagraphProperties4.Append(indentation46);

            level4.Append(startNumberingValue4);
            level4.Append(numberingFormat4);
            level4.Append(levelText4);
            level4.Append(levelJustification4);
            level4.Append(previousParagraphProperties4);

            Level level5 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue5 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat5 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText5 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification5 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties5 = new PreviousParagraphProperties();

            Tabs tabs16 = new Tabs();
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Number, Position = 3360 };

            tabs16.Append(tabStop19);
            Indentation indentation47 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties5.Append(tabs16);
            previousParagraphProperties5.Append(indentation47);

            level5.Append(startNumberingValue5);
            level5.Append(numberingFormat5);
            level5.Append(levelText5);
            level5.Append(levelJustification5);
            level5.Append(previousParagraphProperties5);

            Level level6 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue6 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat6 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText6 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification6 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties6 = new PreviousParagraphProperties();

            Tabs tabs17 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Number, Position = 3840 };

            tabs17.Append(tabStop20);
            Indentation indentation48 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties6.Append(tabs17);
            previousParagraphProperties6.Append(indentation48);

            level6.Append(startNumberingValue6);
            level6.Append(numberingFormat6);
            level6.Append(levelText6);
            level6.Append(levelJustification6);
            level6.Append(previousParagraphProperties6);

            Level level7 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue7 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat7 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText7 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification7 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties7 = new PreviousParagraphProperties();

            Tabs tabs18 = new Tabs();
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Number, Position = 4320 };

            tabs18.Append(tabStop21);
            Indentation indentation49 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties7.Append(tabs18);
            previousParagraphProperties7.Append(indentation49);

            level7.Append(startNumberingValue7);
            level7.Append(numberingFormat7);
            level7.Append(levelText7);
            level7.Append(levelJustification7);
            level7.Append(previousParagraphProperties7);

            Level level8 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue8 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat8 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText8 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification8 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties8 = new PreviousParagraphProperties();

            Tabs tabs19 = new Tabs();
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Number, Position = 4800 };

            tabs19.Append(tabStop22);
            Indentation indentation50 = new Indentation() { Start = "4800", Hanging = "480" };

            previousParagraphProperties8.Append(tabs19);
            previousParagraphProperties8.Append(indentation50);

            level8.Append(startNumberingValue8);
            level8.Append(numberingFormat8);
            level8.Append(levelText8);
            level8.Append(levelJustification8);
            level8.Append(previousParagraphProperties8);

            Level level9 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue9 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat9 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText9 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification9 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties9 = new PreviousParagraphProperties();

            Tabs tabs20 = new Tabs();
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Number, Position = 5280 };

            tabs20.Append(tabStop23);
            Indentation indentation51 = new Indentation() { Start = "5280", Hanging = "480" };

            previousParagraphProperties9.Append(tabs20);
            previousParagraphProperties9.Append(indentation51);

            level9.Append(startNumberingValue9);
            level9.Append(numberingFormat9);
            level9.Append(levelText9);
            level9.Append(levelJustification9);
            level9.Append(previousParagraphProperties9);

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

            AbstractNum abstractNum2 = new AbstractNum() { AbstractNumberId = 1 };
            abstractNum2.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid2 = new Nsid() { Val = "0D104B72" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "24FA0A5A" };

            Level level10 = new Level() { LevelIndex = 0, TemplateCode = "04090015" };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText10 = new LevelText() { Val = "%1、" };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();
            Indentation indentation52 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties10.Append(indentation52);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);

            Level level11 = new Level() { LevelIndex = 1, TemplateCode = "164A8326" };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText11 = new LevelText() { Val = "(%2)" };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();
            Indentation indentation53 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties11.Append(indentation53);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts409 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold284 = new Bold() { Val = false };
            Italic italic1 = new Italic() { Val = false };

            numberingSymbolRunProperties2.Append(runFonts409);
            numberingSymbolRunProperties2.Append(bold284);
            numberingSymbolRunProperties2.Append(italic1);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);
            level11.Append(numberingSymbolRunProperties2);

            Level level12 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText12 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();
            Indentation indentation54 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties12.Append(indentation54);

            level12.Append(startNumberingValue12);
            level12.Append(numberingFormat12);
            level12.Append(levelText12);
            level12.Append(levelJustification12);
            level12.Append(previousParagraphProperties12);

            Level level13 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue13 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat13 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText13 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification13 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties13 = new PreviousParagraphProperties();
            Indentation indentation55 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties13.Append(indentation55);

            level13.Append(startNumberingValue13);
            level13.Append(numberingFormat13);
            level13.Append(levelText13);
            level13.Append(levelJustification13);
            level13.Append(previousParagraphProperties13);

            Level level14 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue14 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat14 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText14 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification14 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties14 = new PreviousParagraphProperties();
            Indentation indentation56 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties14.Append(indentation56);

            level14.Append(startNumberingValue14);
            level14.Append(numberingFormat14);
            level14.Append(levelText14);
            level14.Append(levelJustification14);
            level14.Append(previousParagraphProperties14);

            Level level15 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue15 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat15 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText15 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification15 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties15 = new PreviousParagraphProperties();
            Indentation indentation57 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties15.Append(indentation57);

            level15.Append(startNumberingValue15);
            level15.Append(numberingFormat15);
            level15.Append(levelText15);
            level15.Append(levelJustification15);
            level15.Append(previousParagraphProperties15);

            Level level16 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue16 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat16 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText16 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification16 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties16 = new PreviousParagraphProperties();
            Indentation indentation58 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties16.Append(indentation58);

            level16.Append(startNumberingValue16);
            level16.Append(numberingFormat16);
            level16.Append(levelText16);
            level16.Append(levelJustification16);
            level16.Append(previousParagraphProperties16);

            Level level17 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue17 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat17 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText17 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification17 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties17 = new PreviousParagraphProperties();
            Indentation indentation59 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties17.Append(indentation59);

            level17.Append(startNumberingValue17);
            level17.Append(numberingFormat17);
            level17.Append(levelText17);
            level17.Append(levelJustification17);
            level17.Append(previousParagraphProperties17);

            Level level18 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue18 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat18 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText18 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification18 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties18 = new PreviousParagraphProperties();
            Indentation indentation60 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties18.Append(indentation60);

            level18.Append(startNumberingValue18);
            level18.Append(numberingFormat18);
            level18.Append(levelText18);
            level18.Append(levelJustification18);
            level18.Append(previousParagraphProperties18);

            abstractNum2.Append(nsid2);
            abstractNum2.Append(multiLevelType2);
            abstractNum2.Append(templateCode2);
            abstractNum2.Append(level10);
            abstractNum2.Append(level11);
            abstractNum2.Append(level12);
            abstractNum2.Append(level13);
            abstractNum2.Append(level14);
            abstractNum2.Append(level15);
            abstractNum2.Append(level16);
            abstractNum2.Append(level17);
            abstractNum2.Append(level18);

            AbstractNum abstractNum3 = new AbstractNum() { AbstractNumberId = 2 };
            abstractNum3.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid3 = new Nsid() { Val = "0D330787" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "F13E9CF0" };

            Level level19 = new Level() { LevelIndex = 0, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue19 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat19 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText19 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification19 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties19 = new PreviousParagraphProperties();
            Indentation indentation61 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties19.Append(indentation61);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts410 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties3.Append(runFonts410);

            level19.Append(startNumberingValue19);
            level19.Append(numberingFormat19);
            level19.Append(levelText19);
            level19.Append(levelJustification19);
            level19.Append(previousParagraphProperties19);
            level19.Append(numberingSymbolRunProperties3);

            Level level20 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue20 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat20 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText20 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification20 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties20 = new PreviousParagraphProperties();
            Indentation indentation62 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties20.Append(indentation62);

            level20.Append(startNumberingValue20);
            level20.Append(numberingFormat20);
            level20.Append(levelText20);
            level20.Append(levelJustification20);
            level20.Append(previousParagraphProperties20);

            Level level21 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue21 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat21 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText21 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification21 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties21 = new PreviousParagraphProperties();
            Indentation indentation63 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties21.Append(indentation63);

            level21.Append(startNumberingValue21);
            level21.Append(numberingFormat21);
            level21.Append(levelText21);
            level21.Append(levelJustification21);
            level21.Append(previousParagraphProperties21);

            Level level22 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue22 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat22 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText22 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification22 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties22 = new PreviousParagraphProperties();
            Indentation indentation64 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties22.Append(indentation64);

            level22.Append(startNumberingValue22);
            level22.Append(numberingFormat22);
            level22.Append(levelText22);
            level22.Append(levelJustification22);
            level22.Append(previousParagraphProperties22);

            Level level23 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue23 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat23 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText23 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification23 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties23 = new PreviousParagraphProperties();
            Indentation indentation65 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties23.Append(indentation65);

            level23.Append(startNumberingValue23);
            level23.Append(numberingFormat23);
            level23.Append(levelText23);
            level23.Append(levelJustification23);
            level23.Append(previousParagraphProperties23);

            Level level24 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue24 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat24 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText24 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification24 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties24 = new PreviousParagraphProperties();
            Indentation indentation66 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties24.Append(indentation66);

            level24.Append(startNumberingValue24);
            level24.Append(numberingFormat24);
            level24.Append(levelText24);
            level24.Append(levelJustification24);
            level24.Append(previousParagraphProperties24);

            Level level25 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue25 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat25 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText25 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification25 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties25 = new PreviousParagraphProperties();
            Indentation indentation67 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties25.Append(indentation67);

            level25.Append(startNumberingValue25);
            level25.Append(numberingFormat25);
            level25.Append(levelText25);
            level25.Append(levelJustification25);
            level25.Append(previousParagraphProperties25);

            Level level26 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue26 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat26 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText26 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification26 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties26 = new PreviousParagraphProperties();
            Indentation indentation68 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties26.Append(indentation68);

            level26.Append(startNumberingValue26);
            level26.Append(numberingFormat26);
            level26.Append(levelText26);
            level26.Append(levelJustification26);
            level26.Append(previousParagraphProperties26);

            Level level27 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue27 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat27 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText27 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification27 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties27 = new PreviousParagraphProperties();
            Indentation indentation69 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties27.Append(indentation69);

            level27.Append(startNumberingValue27);
            level27.Append(numberingFormat27);
            level27.Append(levelText27);
            level27.Append(levelJustification27);
            level27.Append(previousParagraphProperties27);

            abstractNum3.Append(nsid3);
            abstractNum3.Append(multiLevelType3);
            abstractNum3.Append(templateCode3);
            abstractNum3.Append(level19);
            abstractNum3.Append(level20);
            abstractNum3.Append(level21);
            abstractNum3.Append(level22);
            abstractNum3.Append(level23);
            abstractNum3.Append(level24);
            abstractNum3.Append(level25);
            abstractNum3.Append(level26);
            abstractNum3.Append(level27);

            AbstractNum abstractNum4 = new AbstractNum() { AbstractNumberId = 3 };
            abstractNum4.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid4 = new Nsid() { Val = "0D8D07FD" };
            MultiLevelType multiLevelType4 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode4 = new TemplateCode() { Val = "5144F146" };

            Level level28 = new Level() { LevelIndex = 0, TemplateCode = "04090013" };
            StartNumberingValue startNumberingValue28 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat28 = new NumberingFormat() { Val = NumberFormatValues.UpperRoman };
            LevelText levelText28 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification28 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties28 = new PreviousParagraphProperties();
            Indentation indentation70 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties28.Append(indentation70);

            level28.Append(startNumberingValue28);
            level28.Append(numberingFormat28);
            level28.Append(levelText28);
            level28.Append(levelJustification28);
            level28.Append(previousParagraphProperties28);

            Level level29 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue29 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat29 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText29 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification29 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties29 = new PreviousParagraphProperties();
            Indentation indentation71 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties29.Append(indentation71);

            level29.Append(startNumberingValue29);
            level29.Append(numberingFormat29);
            level29.Append(levelText29);
            level29.Append(levelJustification29);
            level29.Append(previousParagraphProperties29);

            Level level30 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue30 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat30 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText30 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification30 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties30 = new PreviousParagraphProperties();
            Indentation indentation72 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties30.Append(indentation72);

            level30.Append(startNumberingValue30);
            level30.Append(numberingFormat30);
            level30.Append(levelText30);
            level30.Append(levelJustification30);
            level30.Append(previousParagraphProperties30);

            Level level31 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue31 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat31 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText31 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification31 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties31 = new PreviousParagraphProperties();
            Indentation indentation73 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties31.Append(indentation73);

            level31.Append(startNumberingValue31);
            level31.Append(numberingFormat31);
            level31.Append(levelText31);
            level31.Append(levelJustification31);
            level31.Append(previousParagraphProperties31);

            Level level32 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue32 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat32 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText32 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification32 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties32 = new PreviousParagraphProperties();
            Indentation indentation74 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties32.Append(indentation74);

            level32.Append(startNumberingValue32);
            level32.Append(numberingFormat32);
            level32.Append(levelText32);
            level32.Append(levelJustification32);
            level32.Append(previousParagraphProperties32);

            Level level33 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue33 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat33 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText33 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification33 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties33 = new PreviousParagraphProperties();
            Indentation indentation75 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties33.Append(indentation75);

            level33.Append(startNumberingValue33);
            level33.Append(numberingFormat33);
            level33.Append(levelText33);
            level33.Append(levelJustification33);
            level33.Append(previousParagraphProperties33);

            Level level34 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue34 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat34 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText34 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification34 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties34 = new PreviousParagraphProperties();
            Indentation indentation76 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties34.Append(indentation76);

            level34.Append(startNumberingValue34);
            level34.Append(numberingFormat34);
            level34.Append(levelText34);
            level34.Append(levelJustification34);
            level34.Append(previousParagraphProperties34);

            Level level35 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue35 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat35 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText35 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification35 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties35 = new PreviousParagraphProperties();
            Indentation indentation77 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties35.Append(indentation77);

            level35.Append(startNumberingValue35);
            level35.Append(numberingFormat35);
            level35.Append(levelText35);
            level35.Append(levelJustification35);
            level35.Append(previousParagraphProperties35);

            Level level36 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue36 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat36 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText36 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification36 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties36 = new PreviousParagraphProperties();
            Indentation indentation78 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties36.Append(indentation78);

            level36.Append(startNumberingValue36);
            level36.Append(numberingFormat36);
            level36.Append(levelText36);
            level36.Append(levelJustification36);
            level36.Append(previousParagraphProperties36);

            abstractNum4.Append(nsid4);
            abstractNum4.Append(multiLevelType4);
            abstractNum4.Append(templateCode4);
            abstractNum4.Append(level28);
            abstractNum4.Append(level29);
            abstractNum4.Append(level30);
            abstractNum4.Append(level31);
            abstractNum4.Append(level32);
            abstractNum4.Append(level33);
            abstractNum4.Append(level34);
            abstractNum4.Append(level35);
            abstractNum4.Append(level36);

            AbstractNum abstractNum5 = new AbstractNum() { AbstractNumberId = 4 };
            abstractNum5.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid5 = new Nsid() { Val = "1B914155" };
            MultiLevelType multiLevelType5 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode5 = new TemplateCode() { Val = "2286B7EA" };

            Level level37 = new Level() { LevelIndex = 0, TemplateCode = "04090001" };
            StartNumberingValue startNumberingValue37 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat37 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText37 = new LevelText() { Val = "l" };
            LevelJustification levelJustification37 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties37 = new PreviousParagraphProperties();
            Indentation indentation79 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties37.Append(indentation79);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts411 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties4.Append(runFonts411);

            level37.Append(startNumberingValue37);
            level37.Append(numberingFormat37);
            level37.Append(levelText37);
            level37.Append(levelJustification37);
            level37.Append(previousParagraphProperties37);
            level37.Append(numberingSymbolRunProperties4);

            Level level38 = new Level() { LevelIndex = 1, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue38 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat38 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText38 = new LevelText() { Val = "n" };
            LevelJustification levelJustification38 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties38 = new PreviousParagraphProperties();
            Indentation indentation80 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties38.Append(indentation80);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts412 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties5.Append(runFonts412);

            level38.Append(startNumberingValue38);
            level38.Append(numberingFormat38);
            level38.Append(levelText38);
            level38.Append(levelJustification38);
            level38.Append(previousParagraphProperties38);
            level38.Append(numberingSymbolRunProperties5);

            Level level39 = new Level() { LevelIndex = 2, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue39 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat39 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText39 = new LevelText() { Val = "u" };
            LevelJustification levelJustification39 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties39 = new PreviousParagraphProperties();
            Indentation indentation81 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties39.Append(indentation81);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts413 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties6.Append(runFonts413);

            level39.Append(startNumberingValue39);
            level39.Append(numberingFormat39);
            level39.Append(levelText39);
            level39.Append(levelJustification39);
            level39.Append(previousParagraphProperties39);
            level39.Append(numberingSymbolRunProperties6);

            Level level40 = new Level() { LevelIndex = 3, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue40 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat40 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText40 = new LevelText() { Val = "l" };
            LevelJustification levelJustification40 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties40 = new PreviousParagraphProperties();
            Indentation indentation82 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties40.Append(indentation82);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts414 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties7.Append(runFonts414);

            level40.Append(startNumberingValue40);
            level40.Append(numberingFormat40);
            level40.Append(levelText40);
            level40.Append(levelJustification40);
            level40.Append(previousParagraphProperties40);
            level40.Append(numberingSymbolRunProperties7);

            Level level41 = new Level() { LevelIndex = 4, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue41 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat41 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText41 = new LevelText() { Val = "n" };
            LevelJustification levelJustification41 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties41 = new PreviousParagraphProperties();
            Indentation indentation83 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties41.Append(indentation83);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts415 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties8.Append(runFonts415);

            level41.Append(startNumberingValue41);
            level41.Append(numberingFormat41);
            level41.Append(levelText41);
            level41.Append(levelJustification41);
            level41.Append(previousParagraphProperties41);
            level41.Append(numberingSymbolRunProperties8);

            Level level42 = new Level() { LevelIndex = 5, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue42 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat42 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText42 = new LevelText() { Val = "u" };
            LevelJustification levelJustification42 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties42 = new PreviousParagraphProperties();
            Indentation indentation84 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties42.Append(indentation84);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts416 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties9.Append(runFonts416);

            level42.Append(startNumberingValue42);
            level42.Append(numberingFormat42);
            level42.Append(levelText42);
            level42.Append(levelJustification42);
            level42.Append(previousParagraphProperties42);
            level42.Append(numberingSymbolRunProperties9);

            Level level43 = new Level() { LevelIndex = 6, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue43 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat43 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText43 = new LevelText() { Val = "l" };
            LevelJustification levelJustification43 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties43 = new PreviousParagraphProperties();
            Indentation indentation85 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties43.Append(indentation85);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts417 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties10.Append(runFonts417);

            level43.Append(startNumberingValue43);
            level43.Append(numberingFormat43);
            level43.Append(levelText43);
            level43.Append(levelJustification43);
            level43.Append(previousParagraphProperties43);
            level43.Append(numberingSymbolRunProperties10);

            Level level44 = new Level() { LevelIndex = 7, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue44 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat44 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText44 = new LevelText() { Val = "n" };
            LevelJustification levelJustification44 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties44 = new PreviousParagraphProperties();
            Indentation indentation86 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties44.Append(indentation86);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts418 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties11.Append(runFonts418);

            level44.Append(startNumberingValue44);
            level44.Append(numberingFormat44);
            level44.Append(levelText44);
            level44.Append(levelJustification44);
            level44.Append(previousParagraphProperties44);
            level44.Append(numberingSymbolRunProperties11);

            Level level45 = new Level() { LevelIndex = 8, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue45 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat45 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText45 = new LevelText() { Val = "u" };
            LevelJustification levelJustification45 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties45 = new PreviousParagraphProperties();
            Indentation indentation87 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties45.Append(indentation87);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts419 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties12.Append(runFonts419);

            level45.Append(startNumberingValue45);
            level45.Append(numberingFormat45);
            level45.Append(levelText45);
            level45.Append(levelJustification45);
            level45.Append(previousParagraphProperties45);
            level45.Append(numberingSymbolRunProperties12);

            abstractNum5.Append(nsid5);
            abstractNum5.Append(multiLevelType5);
            abstractNum5.Append(templateCode5);
            abstractNum5.Append(level37);
            abstractNum5.Append(level38);
            abstractNum5.Append(level39);
            abstractNum5.Append(level40);
            abstractNum5.Append(level41);
            abstractNum5.Append(level42);
            abstractNum5.Append(level43);
            abstractNum5.Append(level44);
            abstractNum5.Append(level45);

            AbstractNum abstractNum6 = new AbstractNum() { AbstractNumberId = 5 };
            abstractNum6.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid6 = new Nsid() { Val = "20EE0C91" };
            MultiLevelType multiLevelType6 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode6 = new TemplateCode() { Val = "6FA82208" };

            Level level46 = new Level() { LevelIndex = 0, TemplateCode = "04090017" };
            StartNumberingValue startNumberingValue46 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat46 = new NumberingFormat() { Val = NumberFormatValues.IdeographLegalTraditional };
            LevelText levelText46 = new LevelText() { Val = "%1、" };
            LevelJustification levelJustification46 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties46 = new PreviousParagraphProperties();

            Tabs tabs21 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Number, Position = 482 };

            tabs21.Append(tabStop24);
            Indentation indentation88 = new Indentation() { Start = "482", Hanging = "480" };

            previousParagraphProperties46.Append(tabs21);
            previousParagraphProperties46.Append(indentation88);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts420 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties13.Append(runFonts420);

            level46.Append(startNumberingValue46);
            level46.Append(numberingFormat46);
            level46.Append(levelText46);
            level46.Append(levelJustification46);
            level46.Append(previousParagraphProperties46);
            level46.Append(numberingSymbolRunProperties13);

            Level level47 = new Level() { LevelIndex = 1, TemplateCode = "738AFC8E" };
            StartNumberingValue startNumberingValue47 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat47 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText47 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification47 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties47 = new PreviousParagraphProperties();

            Tabs tabs22 = new Tabs();
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Number, Position = 1082 };

            tabs22.Append(tabStop25);
            Indentation indentation89 = new Indentation() { Start = "1082", Hanging = "600" };

            previousParagraphProperties47.Append(tabs22);
            previousParagraphProperties47.Append(indentation89);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts421 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Spacing spacing1 = new Spacing() { Val = -10 };
            Position position1 = new Position() { Val = "0" };

            numberingSymbolRunProperties14.Append(runFonts421);
            numberingSymbolRunProperties14.Append(spacing1);
            numberingSymbolRunProperties14.Append(position1);

            level47.Append(startNumberingValue47);
            level47.Append(numberingFormat47);
            level47.Append(levelText47);
            level47.Append(levelJustification47);
            level47.Append(previousParagraphProperties47);
            level47.Append(numberingSymbolRunProperties14);

            Level level48 = new Level() { LevelIndex = 2, TemplateCode = "7AD827AC" };
            StartNumberingValue startNumberingValue48 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat48 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText48 = new LevelText() { Val = "(%3)" };
            LevelJustification levelJustification48 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties48 = new PreviousParagraphProperties();

            Tabs tabs23 = new Tabs();
            TabStop tabStop26 = new TabStop() { Val = TabStopValues.Number, Position = 1701 };

            tabs23.Append(tabStop26);
            Indentation indentation90 = new Indentation() { Start = "1701", Hanging = "624" };

            previousParagraphProperties48.Append(tabs23);
            previousParagraphProperties48.Append(indentation90);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts422 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, EastAsia = "標楷體" };
            Bold bold285 = new Bold() { Val = false };
            Italic italic2 = new Italic() { Val = false };
            Spacing spacing2 = new Spacing() { Val = -10 };
            Position position2 = new Position() { Val = "0" };

            numberingSymbolRunProperties15.Append(runFonts422);
            numberingSymbolRunProperties15.Append(bold285);
            numberingSymbolRunProperties15.Append(italic2);
            numberingSymbolRunProperties15.Append(spacing2);
            numberingSymbolRunProperties15.Append(position2);

            level48.Append(startNumberingValue48);
            level48.Append(numberingFormat48);
            level48.Append(levelText48);
            level48.Append(levelJustification48);
            level48.Append(previousParagraphProperties48);
            level48.Append(numberingSymbolRunProperties15);

            Level level49 = new Level() { LevelIndex = 3, TemplateCode = "6814211C" };
            StartNumberingValue startNumberingValue49 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat49 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText49 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification49 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties49 = new PreviousParagraphProperties();

            Tabs tabs24 = new Tabs();
            TabStop tabStop27 = new TabStop() { Val = TabStopValues.Number, Position = 2041 };

            tabs24.Append(tabStop27);
            Indentation indentation91 = new Indentation() { Start = "2041", Hanging = "340" };

            previousParagraphProperties49.Append(tabs24);
            previousParagraphProperties49.Append(indentation91);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts423 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties16.Append(runFonts423);

            level49.Append(startNumberingValue49);
            level49.Append(numberingFormat49);
            level49.Append(levelText49);
            level49.Append(levelJustification49);
            level49.Append(previousParagraphProperties49);
            level49.Append(numberingSymbolRunProperties16);

            Level level50 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue50 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat50 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText50 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification50 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties50 = new PreviousParagraphProperties();

            Tabs tabs25 = new Tabs();
            TabStop tabStop28 = new TabStop() { Val = TabStopValues.Number, Position = 2402 };

            tabs25.Append(tabStop28);
            Indentation indentation92 = new Indentation() { Start = "2402", Hanging = "480" };

            previousParagraphProperties50.Append(tabs25);
            previousParagraphProperties50.Append(indentation92);

            level50.Append(startNumberingValue50);
            level50.Append(numberingFormat50);
            level50.Append(levelText50);
            level50.Append(levelJustification50);
            level50.Append(previousParagraphProperties50);

            Level level51 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue51 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat51 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText51 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification51 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties51 = new PreviousParagraphProperties();

            Tabs tabs26 = new Tabs();
            TabStop tabStop29 = new TabStop() { Val = TabStopValues.Number, Position = 2882 };

            tabs26.Append(tabStop29);
            Indentation indentation93 = new Indentation() { Start = "2882", Hanging = "480" };

            previousParagraphProperties51.Append(tabs26);
            previousParagraphProperties51.Append(indentation93);

            level51.Append(startNumberingValue51);
            level51.Append(numberingFormat51);
            level51.Append(levelText51);
            level51.Append(levelJustification51);
            level51.Append(previousParagraphProperties51);

            Level level52 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue52 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat52 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText52 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification52 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties52 = new PreviousParagraphProperties();

            Tabs tabs27 = new Tabs();
            TabStop tabStop30 = new TabStop() { Val = TabStopValues.Number, Position = 3362 };

            tabs27.Append(tabStop30);
            Indentation indentation94 = new Indentation() { Start = "3362", Hanging = "480" };

            previousParagraphProperties52.Append(tabs27);
            previousParagraphProperties52.Append(indentation94);

            level52.Append(startNumberingValue52);
            level52.Append(numberingFormat52);
            level52.Append(levelText52);
            level52.Append(levelJustification52);
            level52.Append(previousParagraphProperties52);

            Level level53 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue53 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat53 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText53 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification53 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties53 = new PreviousParagraphProperties();

            Tabs tabs28 = new Tabs();
            TabStop tabStop31 = new TabStop() { Val = TabStopValues.Number, Position = 3842 };

            tabs28.Append(tabStop31);
            Indentation indentation95 = new Indentation() { Start = "3842", Hanging = "480" };

            previousParagraphProperties53.Append(tabs28);
            previousParagraphProperties53.Append(indentation95);

            level53.Append(startNumberingValue53);
            level53.Append(numberingFormat53);
            level53.Append(levelText53);
            level53.Append(levelJustification53);
            level53.Append(previousParagraphProperties53);

            Level level54 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue54 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat54 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText54 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification54 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties54 = new PreviousParagraphProperties();

            Tabs tabs29 = new Tabs();
            TabStop tabStop32 = new TabStop() { Val = TabStopValues.Number, Position = 4322 };

            tabs29.Append(tabStop32);
            Indentation indentation96 = new Indentation() { Start = "4322", Hanging = "480" };

            previousParagraphProperties54.Append(tabs29);
            previousParagraphProperties54.Append(indentation96);

            level54.Append(startNumberingValue54);
            level54.Append(numberingFormat54);
            level54.Append(levelText54);
            level54.Append(levelJustification54);
            level54.Append(previousParagraphProperties54);

            abstractNum6.Append(nsid6);
            abstractNum6.Append(multiLevelType6);
            abstractNum6.Append(templateCode6);
            abstractNum6.Append(level46);
            abstractNum6.Append(level47);
            abstractNum6.Append(level48);
            abstractNum6.Append(level49);
            abstractNum6.Append(level50);
            abstractNum6.Append(level51);
            abstractNum6.Append(level52);
            abstractNum6.Append(level53);
            abstractNum6.Append(level54);

            AbstractNum abstractNum7 = new AbstractNum() { AbstractNumberId = 6 };
            abstractNum7.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid7 = new Nsid() { Val = "2DF2620B" };
            MultiLevelType multiLevelType7 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode7 = new TemplateCode() { Val = "93440D90" };

            Level level55 = new Level() { LevelIndex = 0, TemplateCode = "AA78403C" };
            StartNumberingValue startNumberingValue55 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat55 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText55 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification55 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties55 = new PreviousParagraphProperties();
            Indentation indentation97 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties55.Append(indentation97);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts424 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties17.Append(runFonts424);

            level55.Append(startNumberingValue55);
            level55.Append(numberingFormat55);
            level55.Append(levelText55);
            level55.Append(levelJustification55);
            level55.Append(previousParagraphProperties55);
            level55.Append(numberingSymbolRunProperties17);

            Level level56 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue56 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat56 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText56 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification56 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties56 = new PreviousParagraphProperties();
            Indentation indentation98 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties56.Append(indentation98);

            level56.Append(startNumberingValue56);
            level56.Append(numberingFormat56);
            level56.Append(levelText56);
            level56.Append(levelJustification56);
            level56.Append(previousParagraphProperties56);

            Level level57 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue57 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat57 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText57 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification57 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties57 = new PreviousParagraphProperties();
            Indentation indentation99 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties57.Append(indentation99);

            level57.Append(startNumberingValue57);
            level57.Append(numberingFormat57);
            level57.Append(levelText57);
            level57.Append(levelJustification57);
            level57.Append(previousParagraphProperties57);

            Level level58 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue58 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat58 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText58 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification58 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties58 = new PreviousParagraphProperties();
            Indentation indentation100 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties58.Append(indentation100);

            level58.Append(startNumberingValue58);
            level58.Append(numberingFormat58);
            level58.Append(levelText58);
            level58.Append(levelJustification58);
            level58.Append(previousParagraphProperties58);

            Level level59 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue59 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat59 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText59 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification59 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties59 = new PreviousParagraphProperties();
            Indentation indentation101 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties59.Append(indentation101);

            level59.Append(startNumberingValue59);
            level59.Append(numberingFormat59);
            level59.Append(levelText59);
            level59.Append(levelJustification59);
            level59.Append(previousParagraphProperties59);

            Level level60 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue60 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat60 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText60 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification60 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties60 = new PreviousParagraphProperties();
            Indentation indentation102 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties60.Append(indentation102);

            level60.Append(startNumberingValue60);
            level60.Append(numberingFormat60);
            level60.Append(levelText60);
            level60.Append(levelJustification60);
            level60.Append(previousParagraphProperties60);

            Level level61 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue61 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat61 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText61 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification61 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties61 = new PreviousParagraphProperties();
            Indentation indentation103 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties61.Append(indentation103);

            level61.Append(startNumberingValue61);
            level61.Append(numberingFormat61);
            level61.Append(levelText61);
            level61.Append(levelJustification61);
            level61.Append(previousParagraphProperties61);

            Level level62 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue62 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat62 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText62 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification62 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties62 = new PreviousParagraphProperties();
            Indentation indentation104 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties62.Append(indentation104);

            level62.Append(startNumberingValue62);
            level62.Append(numberingFormat62);
            level62.Append(levelText62);
            level62.Append(levelJustification62);
            level62.Append(previousParagraphProperties62);

            Level level63 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue63 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat63 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText63 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification63 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties63 = new PreviousParagraphProperties();
            Indentation indentation105 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties63.Append(indentation105);

            level63.Append(startNumberingValue63);
            level63.Append(numberingFormat63);
            level63.Append(levelText63);
            level63.Append(levelJustification63);
            level63.Append(previousParagraphProperties63);

            abstractNum7.Append(nsid7);
            abstractNum7.Append(multiLevelType7);
            abstractNum7.Append(templateCode7);
            abstractNum7.Append(level55);
            abstractNum7.Append(level56);
            abstractNum7.Append(level57);
            abstractNum7.Append(level58);
            abstractNum7.Append(level59);
            abstractNum7.Append(level60);
            abstractNum7.Append(level61);
            abstractNum7.Append(level62);
            abstractNum7.Append(level63);

            AbstractNum abstractNum8 = new AbstractNum() { AbstractNumberId = 7 };
            abstractNum8.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid8 = new Nsid() { Val = "3AD5471D" };
            MultiLevelType multiLevelType8 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode8 = new TemplateCode() { Val = "D58AC6F4" };

            Level level64 = new Level() { LevelIndex = 0, TemplateCode = "1046A65E" };
            StartNumberingValue startNumberingValue64 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat64 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText64 = new LevelText() { Val = "(%1)" };
            LevelJustification levelJustification64 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties64 = new PreviousParagraphProperties();

            Tabs tabs30 = new Tabs();
            TabStop tabStop33 = new TabStop() { Val = TabStopValues.Number, Position = 1284 };

            tabs30.Append(tabStop33);
            Indentation indentation106 = new Indentation() { Start = "1284", Hanging = "720" };

            previousParagraphProperties64.Append(tabs30);
            previousParagraphProperties64.Append(indentation106);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts425 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties18.Append(runFonts425);

            level64.Append(startNumberingValue64);
            level64.Append(numberingFormat64);
            level64.Append(levelText64);
            level64.Append(levelJustification64);
            level64.Append(previousParagraphProperties64);
            level64.Append(numberingSymbolRunProperties18);

            Level level65 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue65 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat65 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText65 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification65 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties65 = new PreviousParagraphProperties();

            Tabs tabs31 = new Tabs();
            TabStop tabStop34 = new TabStop() { Val = TabStopValues.Number, Position = 1524 };

            tabs31.Append(tabStop34);
            Indentation indentation107 = new Indentation() { Start = "1524", Hanging = "480" };

            previousParagraphProperties65.Append(tabs31);
            previousParagraphProperties65.Append(indentation107);

            level65.Append(startNumberingValue65);
            level65.Append(numberingFormat65);
            level65.Append(levelText65);
            level65.Append(levelJustification65);
            level65.Append(previousParagraphProperties65);

            Level level66 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue66 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat66 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText66 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification66 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties66 = new PreviousParagraphProperties();

            Tabs tabs32 = new Tabs();
            TabStop tabStop35 = new TabStop() { Val = TabStopValues.Number, Position = 2004 };

            tabs32.Append(tabStop35);
            Indentation indentation108 = new Indentation() { Start = "2004", Hanging = "480" };

            previousParagraphProperties66.Append(tabs32);
            previousParagraphProperties66.Append(indentation108);

            level66.Append(startNumberingValue66);
            level66.Append(numberingFormat66);
            level66.Append(levelText66);
            level66.Append(levelJustification66);
            level66.Append(previousParagraphProperties66);

            Level level67 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue67 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat67 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText67 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification67 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties67 = new PreviousParagraphProperties();

            Tabs tabs33 = new Tabs();
            TabStop tabStop36 = new TabStop() { Val = TabStopValues.Number, Position = 2484 };

            tabs33.Append(tabStop36);
            Indentation indentation109 = new Indentation() { Start = "2484", Hanging = "480" };

            previousParagraphProperties67.Append(tabs33);
            previousParagraphProperties67.Append(indentation109);

            level67.Append(startNumberingValue67);
            level67.Append(numberingFormat67);
            level67.Append(levelText67);
            level67.Append(levelJustification67);
            level67.Append(previousParagraphProperties67);

            Level level68 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue68 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat68 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText68 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification68 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties68 = new PreviousParagraphProperties();

            Tabs tabs34 = new Tabs();
            TabStop tabStop37 = new TabStop() { Val = TabStopValues.Number, Position = 2964 };

            tabs34.Append(tabStop37);
            Indentation indentation110 = new Indentation() { Start = "2964", Hanging = "480" };

            previousParagraphProperties68.Append(tabs34);
            previousParagraphProperties68.Append(indentation110);

            level68.Append(startNumberingValue68);
            level68.Append(numberingFormat68);
            level68.Append(levelText68);
            level68.Append(levelJustification68);
            level68.Append(previousParagraphProperties68);

            Level level69 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue69 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat69 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText69 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification69 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties69 = new PreviousParagraphProperties();

            Tabs tabs35 = new Tabs();
            TabStop tabStop38 = new TabStop() { Val = TabStopValues.Number, Position = 3444 };

            tabs35.Append(tabStop38);
            Indentation indentation111 = new Indentation() { Start = "3444", Hanging = "480" };

            previousParagraphProperties69.Append(tabs35);
            previousParagraphProperties69.Append(indentation111);

            level69.Append(startNumberingValue69);
            level69.Append(numberingFormat69);
            level69.Append(levelText69);
            level69.Append(levelJustification69);
            level69.Append(previousParagraphProperties69);

            Level level70 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue70 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat70 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText70 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification70 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties70 = new PreviousParagraphProperties();

            Tabs tabs36 = new Tabs();
            TabStop tabStop39 = new TabStop() { Val = TabStopValues.Number, Position = 3924 };

            tabs36.Append(tabStop39);
            Indentation indentation112 = new Indentation() { Start = "3924", Hanging = "480" };

            previousParagraphProperties70.Append(tabs36);
            previousParagraphProperties70.Append(indentation112);

            level70.Append(startNumberingValue70);
            level70.Append(numberingFormat70);
            level70.Append(levelText70);
            level70.Append(levelJustification70);
            level70.Append(previousParagraphProperties70);

            Level level71 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue71 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat71 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText71 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification71 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties71 = new PreviousParagraphProperties();

            Tabs tabs37 = new Tabs();
            TabStop tabStop40 = new TabStop() { Val = TabStopValues.Number, Position = 4404 };

            tabs37.Append(tabStop40);
            Indentation indentation113 = new Indentation() { Start = "4404", Hanging = "480" };

            previousParagraphProperties71.Append(tabs37);
            previousParagraphProperties71.Append(indentation113);

            level71.Append(startNumberingValue71);
            level71.Append(numberingFormat71);
            level71.Append(levelText71);
            level71.Append(levelJustification71);
            level71.Append(previousParagraphProperties71);

            Level level72 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue72 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat72 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText72 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification72 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties72 = new PreviousParagraphProperties();

            Tabs tabs38 = new Tabs();
            TabStop tabStop41 = new TabStop() { Val = TabStopValues.Number, Position = 4884 };

            tabs38.Append(tabStop41);
            Indentation indentation114 = new Indentation() { Start = "4884", Hanging = "480" };

            previousParagraphProperties72.Append(tabs38);
            previousParagraphProperties72.Append(indentation114);

            level72.Append(startNumberingValue72);
            level72.Append(numberingFormat72);
            level72.Append(levelText72);
            level72.Append(levelJustification72);
            level72.Append(previousParagraphProperties72);

            abstractNum8.Append(nsid8);
            abstractNum8.Append(multiLevelType8);
            abstractNum8.Append(templateCode8);
            abstractNum8.Append(level64);
            abstractNum8.Append(level65);
            abstractNum8.Append(level66);
            abstractNum8.Append(level67);
            abstractNum8.Append(level68);
            abstractNum8.Append(level69);
            abstractNum8.Append(level70);
            abstractNum8.Append(level71);
            abstractNum8.Append(level72);

            AbstractNum abstractNum9 = new AbstractNum() { AbstractNumberId = 8 };
            abstractNum9.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid9 = new Nsid() { Val = "41A04F1E" };
            MultiLevelType multiLevelType9 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode9 = new TemplateCode() { Val = "0024B006" };

            Level level73 = new Level() { LevelIndex = 0, TemplateCode = "55D8A5B4" };
            StartNumberingValue startNumberingValue73 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat73 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText73 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification73 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties73 = new PreviousParagraphProperties();
            Indentation indentation115 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties73.Append(indentation115);

            NumberingSymbolRunProperties numberingSymbolRunProperties19 = new NumberingSymbolRunProperties();
            RunFonts runFonts426 = new RunFonts() { Hint = FontTypeHintValues.Default };
            Color color199 = new Color() { Val = "000000" };

            numberingSymbolRunProperties19.Append(runFonts426);
            numberingSymbolRunProperties19.Append(color199);

            level73.Append(startNumberingValue73);
            level73.Append(numberingFormat73);
            level73.Append(levelText73);
            level73.Append(levelJustification73);
            level73.Append(previousParagraphProperties73);
            level73.Append(numberingSymbolRunProperties19);

            Level level74 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue74 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat74 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText74 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification74 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties74 = new PreviousParagraphProperties();
            Indentation indentation116 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties74.Append(indentation116);

            level74.Append(startNumberingValue74);
            level74.Append(numberingFormat74);
            level74.Append(levelText74);
            level74.Append(levelJustification74);
            level74.Append(previousParagraphProperties74);

            Level level75 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue75 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat75 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText75 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification75 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties75 = new PreviousParagraphProperties();
            Indentation indentation117 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties75.Append(indentation117);

            level75.Append(startNumberingValue75);
            level75.Append(numberingFormat75);
            level75.Append(levelText75);
            level75.Append(levelJustification75);
            level75.Append(previousParagraphProperties75);

            Level level76 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue76 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat76 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText76 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification76 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties76 = new PreviousParagraphProperties();
            Indentation indentation118 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties76.Append(indentation118);

            level76.Append(startNumberingValue76);
            level76.Append(numberingFormat76);
            level76.Append(levelText76);
            level76.Append(levelJustification76);
            level76.Append(previousParagraphProperties76);

            Level level77 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue77 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat77 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText77 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification77 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties77 = new PreviousParagraphProperties();
            Indentation indentation119 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties77.Append(indentation119);

            level77.Append(startNumberingValue77);
            level77.Append(numberingFormat77);
            level77.Append(levelText77);
            level77.Append(levelJustification77);
            level77.Append(previousParagraphProperties77);

            Level level78 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue78 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat78 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText78 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification78 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties78 = new PreviousParagraphProperties();
            Indentation indentation120 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties78.Append(indentation120);

            level78.Append(startNumberingValue78);
            level78.Append(numberingFormat78);
            level78.Append(levelText78);
            level78.Append(levelJustification78);
            level78.Append(previousParagraphProperties78);

            Level level79 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue79 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat79 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText79 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification79 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties79 = new PreviousParagraphProperties();
            Indentation indentation121 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties79.Append(indentation121);

            level79.Append(startNumberingValue79);
            level79.Append(numberingFormat79);
            level79.Append(levelText79);
            level79.Append(levelJustification79);
            level79.Append(previousParagraphProperties79);

            Level level80 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue80 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat80 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText80 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification80 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties80 = new PreviousParagraphProperties();
            Indentation indentation122 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties80.Append(indentation122);

            level80.Append(startNumberingValue80);
            level80.Append(numberingFormat80);
            level80.Append(levelText80);
            level80.Append(levelJustification80);
            level80.Append(previousParagraphProperties80);

            Level level81 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue81 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat81 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText81 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification81 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties81 = new PreviousParagraphProperties();
            Indentation indentation123 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties81.Append(indentation123);

            level81.Append(startNumberingValue81);
            level81.Append(numberingFormat81);
            level81.Append(levelText81);
            level81.Append(levelJustification81);
            level81.Append(previousParagraphProperties81);

            abstractNum9.Append(nsid9);
            abstractNum9.Append(multiLevelType9);
            abstractNum9.Append(templateCode9);
            abstractNum9.Append(level73);
            abstractNum9.Append(level74);
            abstractNum9.Append(level75);
            abstractNum9.Append(level76);
            abstractNum9.Append(level77);
            abstractNum9.Append(level78);
            abstractNum9.Append(level79);
            abstractNum9.Append(level80);
            abstractNum9.Append(level81);

            AbstractNum abstractNum10 = new AbstractNum() { AbstractNumberId = 9 };
            abstractNum10.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid10 = new Nsid() { Val = "44176B93" };
            MultiLevelType multiLevelType10 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode10 = new TemplateCode() { Val = "54269EFA" };

            Level level82 = new Level() { LevelIndex = 0, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue82 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat82 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText82 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification82 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties82 = new PreviousParagraphProperties();
            Indentation indentation124 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties82.Append(indentation124);

            level82.Append(startNumberingValue82);
            level82.Append(numberingFormat82);
            level82.Append(levelText82);
            level82.Append(levelJustification82);
            level82.Append(previousParagraphProperties82);

            Level level83 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue83 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat83 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText83 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification83 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties83 = new PreviousParagraphProperties();
            Indentation indentation125 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties83.Append(indentation125);

            level83.Append(startNumberingValue83);
            level83.Append(numberingFormat83);
            level83.Append(levelText83);
            level83.Append(levelJustification83);
            level83.Append(previousParagraphProperties83);

            Level level84 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue84 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat84 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText84 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification84 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties84 = new PreviousParagraphProperties();
            Indentation indentation126 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties84.Append(indentation126);

            level84.Append(startNumberingValue84);
            level84.Append(numberingFormat84);
            level84.Append(levelText84);
            level84.Append(levelJustification84);
            level84.Append(previousParagraphProperties84);

            Level level85 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue85 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat85 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText85 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification85 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties85 = new PreviousParagraphProperties();
            Indentation indentation127 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties85.Append(indentation127);

            level85.Append(startNumberingValue85);
            level85.Append(numberingFormat85);
            level85.Append(levelText85);
            level85.Append(levelJustification85);
            level85.Append(previousParagraphProperties85);

            Level level86 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue86 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat86 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText86 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification86 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties86 = new PreviousParagraphProperties();
            Indentation indentation128 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties86.Append(indentation128);

            level86.Append(startNumberingValue86);
            level86.Append(numberingFormat86);
            level86.Append(levelText86);
            level86.Append(levelJustification86);
            level86.Append(previousParagraphProperties86);

            Level level87 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue87 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat87 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText87 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification87 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties87 = new PreviousParagraphProperties();
            Indentation indentation129 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties87.Append(indentation129);

            level87.Append(startNumberingValue87);
            level87.Append(numberingFormat87);
            level87.Append(levelText87);
            level87.Append(levelJustification87);
            level87.Append(previousParagraphProperties87);

            Level level88 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue88 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat88 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText88 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification88 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties88 = new PreviousParagraphProperties();
            Indentation indentation130 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties88.Append(indentation130);

            level88.Append(startNumberingValue88);
            level88.Append(numberingFormat88);
            level88.Append(levelText88);
            level88.Append(levelJustification88);
            level88.Append(previousParagraphProperties88);

            Level level89 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue89 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat89 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText89 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification89 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties89 = new PreviousParagraphProperties();
            Indentation indentation131 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties89.Append(indentation131);

            level89.Append(startNumberingValue89);
            level89.Append(numberingFormat89);
            level89.Append(levelText89);
            level89.Append(levelJustification89);
            level89.Append(previousParagraphProperties89);

            Level level90 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue90 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat90 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText90 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification90 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties90 = new PreviousParagraphProperties();
            Indentation indentation132 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties90.Append(indentation132);

            level90.Append(startNumberingValue90);
            level90.Append(numberingFormat90);
            level90.Append(levelText90);
            level90.Append(levelJustification90);
            level90.Append(previousParagraphProperties90);

            abstractNum10.Append(nsid10);
            abstractNum10.Append(multiLevelType10);
            abstractNum10.Append(templateCode10);
            abstractNum10.Append(level82);
            abstractNum10.Append(level83);
            abstractNum10.Append(level84);
            abstractNum10.Append(level85);
            abstractNum10.Append(level86);
            abstractNum10.Append(level87);
            abstractNum10.Append(level88);
            abstractNum10.Append(level89);
            abstractNum10.Append(level90);

            AbstractNum abstractNum11 = new AbstractNum() { AbstractNumberId = 10 };
            abstractNum11.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid11 = new Nsid() { Val = "49EC268A" };
            MultiLevelType multiLevelType11 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode11 = new TemplateCode() { Val = "927E4F54" };

            Level level91 = new Level() { LevelIndex = 0, TemplateCode = "A6F0E89E" };
            StartNumberingValue startNumberingValue91 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat91 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText91 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification91 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties91 = new PreviousParagraphProperties();
            Indentation indentation133 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties91.Append(indentation133);

            NumberingSymbolRunProperties numberingSymbolRunProperties20 = new NumberingSymbolRunProperties();
            RunFonts runFonts427 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties20.Append(runFonts427);

            level91.Append(startNumberingValue91);
            level91.Append(numberingFormat91);
            level91.Append(levelText91);
            level91.Append(levelJustification91);
            level91.Append(previousParagraphProperties91);
            level91.Append(numberingSymbolRunProperties20);

            Level level92 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue92 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat92 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText92 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification92 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties92 = new PreviousParagraphProperties();
            Indentation indentation134 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties92.Append(indentation134);

            level92.Append(startNumberingValue92);
            level92.Append(numberingFormat92);
            level92.Append(levelText92);
            level92.Append(levelJustification92);
            level92.Append(previousParagraphProperties92);

            Level level93 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue93 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat93 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText93 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification93 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties93 = new PreviousParagraphProperties();
            Indentation indentation135 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties93.Append(indentation135);

            level93.Append(startNumberingValue93);
            level93.Append(numberingFormat93);
            level93.Append(levelText93);
            level93.Append(levelJustification93);
            level93.Append(previousParagraphProperties93);

            Level level94 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue94 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat94 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText94 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification94 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties94 = new PreviousParagraphProperties();
            Indentation indentation136 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties94.Append(indentation136);

            level94.Append(startNumberingValue94);
            level94.Append(numberingFormat94);
            level94.Append(levelText94);
            level94.Append(levelJustification94);
            level94.Append(previousParagraphProperties94);

            Level level95 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue95 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat95 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText95 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification95 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties95 = new PreviousParagraphProperties();
            Indentation indentation137 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties95.Append(indentation137);

            level95.Append(startNumberingValue95);
            level95.Append(numberingFormat95);
            level95.Append(levelText95);
            level95.Append(levelJustification95);
            level95.Append(previousParagraphProperties95);

            Level level96 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue96 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat96 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText96 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification96 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties96 = new PreviousParagraphProperties();
            Indentation indentation138 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties96.Append(indentation138);

            level96.Append(startNumberingValue96);
            level96.Append(numberingFormat96);
            level96.Append(levelText96);
            level96.Append(levelJustification96);
            level96.Append(previousParagraphProperties96);

            Level level97 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue97 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat97 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText97 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification97 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties97 = new PreviousParagraphProperties();
            Indentation indentation139 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties97.Append(indentation139);

            level97.Append(startNumberingValue97);
            level97.Append(numberingFormat97);
            level97.Append(levelText97);
            level97.Append(levelJustification97);
            level97.Append(previousParagraphProperties97);

            Level level98 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue98 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat98 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText98 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification98 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties98 = new PreviousParagraphProperties();
            Indentation indentation140 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties98.Append(indentation140);

            level98.Append(startNumberingValue98);
            level98.Append(numberingFormat98);
            level98.Append(levelText98);
            level98.Append(levelJustification98);
            level98.Append(previousParagraphProperties98);

            Level level99 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue99 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat99 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText99 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification99 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties99 = new PreviousParagraphProperties();
            Indentation indentation141 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties99.Append(indentation141);

            level99.Append(startNumberingValue99);
            level99.Append(numberingFormat99);
            level99.Append(levelText99);
            level99.Append(levelJustification99);
            level99.Append(previousParagraphProperties99);

            abstractNum11.Append(nsid11);
            abstractNum11.Append(multiLevelType11);
            abstractNum11.Append(templateCode11);
            abstractNum11.Append(level91);
            abstractNum11.Append(level92);
            abstractNum11.Append(level93);
            abstractNum11.Append(level94);
            abstractNum11.Append(level95);
            abstractNum11.Append(level96);
            abstractNum11.Append(level97);
            abstractNum11.Append(level98);
            abstractNum11.Append(level99);

            AbstractNum abstractNum12 = new AbstractNum() { AbstractNumberId = 11 };
            abstractNum12.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid12 = new Nsid() { Val = "61332F21" };
            MultiLevelType multiLevelType12 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode12 = new TemplateCode() { Val = "92F2F87C" };

            Level level100 = new Level() { LevelIndex = 0, TemplateCode = "0409001B" };
            StartNumberingValue startNumberingValue100 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat100 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText100 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification100 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties100 = new PreviousParagraphProperties();

            Tabs tabs39 = new Tabs();
            TabStop tabStop42 = new TabStop() { Val = TabStopValues.Number, Position = 1680 };

            tabs39.Append(tabStop42);
            Indentation indentation142 = new Indentation() { Start = "1680", Hanging = "720" };

            previousParagraphProperties100.Append(tabs39);
            previousParagraphProperties100.Append(indentation142);

            NumberingSymbolRunProperties numberingSymbolRunProperties21 = new NumberingSymbolRunProperties();
            RunFonts runFonts428 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties21.Append(runFonts428);

            level100.Append(startNumberingValue100);
            level100.Append(numberingFormat100);
            level100.Append(levelText100);
            level100.Append(levelJustification100);
            level100.Append(previousParagraphProperties100);
            level100.Append(numberingSymbolRunProperties21);

            Level level101 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue101 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat101 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText101 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification101 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties101 = new PreviousParagraphProperties();

            Tabs tabs40 = new Tabs();
            TabStop tabStop43 = new TabStop() { Val = TabStopValues.Number, Position = 1920 };

            tabs40.Append(tabStop43);
            Indentation indentation143 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties101.Append(tabs40);
            previousParagraphProperties101.Append(indentation143);

            level101.Append(startNumberingValue101);
            level101.Append(numberingFormat101);
            level101.Append(levelText101);
            level101.Append(levelJustification101);
            level101.Append(previousParagraphProperties101);

            Level level102 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue102 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat102 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText102 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification102 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties102 = new PreviousParagraphProperties();

            Tabs tabs41 = new Tabs();
            TabStop tabStop44 = new TabStop() { Val = TabStopValues.Number, Position = 2400 };

            tabs41.Append(tabStop44);
            Indentation indentation144 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties102.Append(tabs41);
            previousParagraphProperties102.Append(indentation144);

            level102.Append(startNumberingValue102);
            level102.Append(numberingFormat102);
            level102.Append(levelText102);
            level102.Append(levelJustification102);
            level102.Append(previousParagraphProperties102);

            Level level103 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue103 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat103 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText103 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification103 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties103 = new PreviousParagraphProperties();

            Tabs tabs42 = new Tabs();
            TabStop tabStop45 = new TabStop() { Val = TabStopValues.Number, Position = 2880 };

            tabs42.Append(tabStop45);
            Indentation indentation145 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties103.Append(tabs42);
            previousParagraphProperties103.Append(indentation145);

            level103.Append(startNumberingValue103);
            level103.Append(numberingFormat103);
            level103.Append(levelText103);
            level103.Append(levelJustification103);
            level103.Append(previousParagraphProperties103);

            Level level104 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue104 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat104 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText104 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification104 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties104 = new PreviousParagraphProperties();

            Tabs tabs43 = new Tabs();
            TabStop tabStop46 = new TabStop() { Val = TabStopValues.Number, Position = 3360 };

            tabs43.Append(tabStop46);
            Indentation indentation146 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties104.Append(tabs43);
            previousParagraphProperties104.Append(indentation146);

            level104.Append(startNumberingValue104);
            level104.Append(numberingFormat104);
            level104.Append(levelText104);
            level104.Append(levelJustification104);
            level104.Append(previousParagraphProperties104);

            Level level105 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue105 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat105 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText105 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification105 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties105 = new PreviousParagraphProperties();

            Tabs tabs44 = new Tabs();
            TabStop tabStop47 = new TabStop() { Val = TabStopValues.Number, Position = 3840 };

            tabs44.Append(tabStop47);
            Indentation indentation147 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties105.Append(tabs44);
            previousParagraphProperties105.Append(indentation147);

            level105.Append(startNumberingValue105);
            level105.Append(numberingFormat105);
            level105.Append(levelText105);
            level105.Append(levelJustification105);
            level105.Append(previousParagraphProperties105);

            Level level106 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue106 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat106 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText106 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification106 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties106 = new PreviousParagraphProperties();

            Tabs tabs45 = new Tabs();
            TabStop tabStop48 = new TabStop() { Val = TabStopValues.Number, Position = 4320 };

            tabs45.Append(tabStop48);
            Indentation indentation148 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties106.Append(tabs45);
            previousParagraphProperties106.Append(indentation148);

            level106.Append(startNumberingValue106);
            level106.Append(numberingFormat106);
            level106.Append(levelText106);
            level106.Append(levelJustification106);
            level106.Append(previousParagraphProperties106);

            Level level107 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue107 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat107 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText107 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification107 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties107 = new PreviousParagraphProperties();

            Tabs tabs46 = new Tabs();
            TabStop tabStop49 = new TabStop() { Val = TabStopValues.Number, Position = 4800 };

            tabs46.Append(tabStop49);
            Indentation indentation149 = new Indentation() { Start = "4800", Hanging = "480" };

            previousParagraphProperties107.Append(tabs46);
            previousParagraphProperties107.Append(indentation149);

            level107.Append(startNumberingValue107);
            level107.Append(numberingFormat107);
            level107.Append(levelText107);
            level107.Append(levelJustification107);
            level107.Append(previousParagraphProperties107);

            Level level108 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue108 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat108 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText108 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification108 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties108 = new PreviousParagraphProperties();

            Tabs tabs47 = new Tabs();
            TabStop tabStop50 = new TabStop() { Val = TabStopValues.Number, Position = 5280 };

            tabs47.Append(tabStop50);
            Indentation indentation150 = new Indentation() { Start = "5280", Hanging = "480" };

            previousParagraphProperties108.Append(tabs47);
            previousParagraphProperties108.Append(indentation150);

            level108.Append(startNumberingValue108);
            level108.Append(numberingFormat108);
            level108.Append(levelText108);
            level108.Append(levelJustification108);
            level108.Append(previousParagraphProperties108);

            abstractNum12.Append(nsid12);
            abstractNum12.Append(multiLevelType12);
            abstractNum12.Append(templateCode12);
            abstractNum12.Append(level100);
            abstractNum12.Append(level101);
            abstractNum12.Append(level102);
            abstractNum12.Append(level103);
            abstractNum12.Append(level104);
            abstractNum12.Append(level105);
            abstractNum12.Append(level106);
            abstractNum12.Append(level107);
            abstractNum12.Append(level108);

            AbstractNum abstractNum13 = new AbstractNum() { AbstractNumberId = 12 };
            abstractNum13.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid13 = new Nsid() { Val = "64403189" };
            MultiLevelType multiLevelType13 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode13 = new TemplateCode() { Val = "0F50D422" };

            Level level109 = new Level() { LevelIndex = 0, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue109 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat109 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText109 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification109 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties109 = new PreviousParagraphProperties();
            Indentation indentation151 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties109.Append(indentation151);

            level109.Append(startNumberingValue109);
            level109.Append(numberingFormat109);
            level109.Append(levelText109);
            level109.Append(levelJustification109);
            level109.Append(previousParagraphProperties109);

            Level level110 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue110 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat110 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText110 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification110 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties110 = new PreviousParagraphProperties();
            Indentation indentation152 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties110.Append(indentation152);

            level110.Append(startNumberingValue110);
            level110.Append(numberingFormat110);
            level110.Append(levelText110);
            level110.Append(levelJustification110);
            level110.Append(previousParagraphProperties110);

            Level level111 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue111 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat111 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText111 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification111 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties111 = new PreviousParagraphProperties();
            Indentation indentation153 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties111.Append(indentation153);

            level111.Append(startNumberingValue111);
            level111.Append(numberingFormat111);
            level111.Append(levelText111);
            level111.Append(levelJustification111);
            level111.Append(previousParagraphProperties111);

            Level level112 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue112 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat112 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText112 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification112 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties112 = new PreviousParagraphProperties();
            Indentation indentation154 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties112.Append(indentation154);

            level112.Append(startNumberingValue112);
            level112.Append(numberingFormat112);
            level112.Append(levelText112);
            level112.Append(levelJustification112);
            level112.Append(previousParagraphProperties112);

            Level level113 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue113 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat113 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText113 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification113 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties113 = new PreviousParagraphProperties();
            Indentation indentation155 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties113.Append(indentation155);

            level113.Append(startNumberingValue113);
            level113.Append(numberingFormat113);
            level113.Append(levelText113);
            level113.Append(levelJustification113);
            level113.Append(previousParagraphProperties113);

            Level level114 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue114 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat114 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText114 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification114 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties114 = new PreviousParagraphProperties();
            Indentation indentation156 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties114.Append(indentation156);

            level114.Append(startNumberingValue114);
            level114.Append(numberingFormat114);
            level114.Append(levelText114);
            level114.Append(levelJustification114);
            level114.Append(previousParagraphProperties114);

            Level level115 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue115 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat115 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText115 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification115 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties115 = new PreviousParagraphProperties();
            Indentation indentation157 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties115.Append(indentation157);

            level115.Append(startNumberingValue115);
            level115.Append(numberingFormat115);
            level115.Append(levelText115);
            level115.Append(levelJustification115);
            level115.Append(previousParagraphProperties115);

            Level level116 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue116 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat116 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText116 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification116 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties116 = new PreviousParagraphProperties();
            Indentation indentation158 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties116.Append(indentation158);

            level116.Append(startNumberingValue116);
            level116.Append(numberingFormat116);
            level116.Append(levelText116);
            level116.Append(levelJustification116);
            level116.Append(previousParagraphProperties116);

            Level level117 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue117 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat117 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText117 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification117 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties117 = new PreviousParagraphProperties();
            Indentation indentation159 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties117.Append(indentation159);

            level117.Append(startNumberingValue117);
            level117.Append(numberingFormat117);
            level117.Append(levelText117);
            level117.Append(levelJustification117);
            level117.Append(previousParagraphProperties117);

            abstractNum13.Append(nsid13);
            abstractNum13.Append(multiLevelType13);
            abstractNum13.Append(templateCode13);
            abstractNum13.Append(level109);
            abstractNum13.Append(level110);
            abstractNum13.Append(level111);
            abstractNum13.Append(level112);
            abstractNum13.Append(level113);
            abstractNum13.Append(level114);
            abstractNum13.Append(level115);
            abstractNum13.Append(level116);
            abstractNum13.Append(level117);

            AbstractNum abstractNum14 = new AbstractNum() { AbstractNumberId = 13 };
            abstractNum14.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid14 = new Nsid() { Val = "67E86D56" };
            MultiLevelType multiLevelType14 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode14 = new TemplateCode() { Val = "5FFE0390" };

            Level level118 = new Level() { LevelIndex = 0, TemplateCode = "04090013" };
            StartNumberingValue startNumberingValue118 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat118 = new NumberingFormat() { Val = NumberFormatValues.UpperRoman };
            LevelText levelText118 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification118 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties118 = new PreviousParagraphProperties();
            Indentation indentation160 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties118.Append(indentation160);

            level118.Append(startNumberingValue118);
            level118.Append(numberingFormat118);
            level118.Append(levelText118);
            level118.Append(levelJustification118);
            level118.Append(previousParagraphProperties118);

            Level level119 = new Level() { LevelIndex = 1, TemplateCode = "04090019" };
            StartNumberingValue startNumberingValue119 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat119 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText119 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification119 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties119 = new PreviousParagraphProperties();
            Indentation indentation161 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties119.Append(indentation161);

            level119.Append(startNumberingValue119);
            level119.Append(numberingFormat119);
            level119.Append(levelText119);
            level119.Append(levelJustification119);
            level119.Append(previousParagraphProperties119);

            Level level120 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue120 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat120 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText120 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification120 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties120 = new PreviousParagraphProperties();
            Indentation indentation162 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties120.Append(indentation162);

            level120.Append(startNumberingValue120);
            level120.Append(numberingFormat120);
            level120.Append(levelText120);
            level120.Append(levelJustification120);
            level120.Append(previousParagraphProperties120);

            Level level121 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue121 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat121 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText121 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification121 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties121 = new PreviousParagraphProperties();
            Indentation indentation163 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties121.Append(indentation163);

            level121.Append(startNumberingValue121);
            level121.Append(numberingFormat121);
            level121.Append(levelText121);
            level121.Append(levelJustification121);
            level121.Append(previousParagraphProperties121);

            Level level122 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue122 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat122 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText122 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification122 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties122 = new PreviousParagraphProperties();
            Indentation indentation164 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties122.Append(indentation164);

            level122.Append(startNumberingValue122);
            level122.Append(numberingFormat122);
            level122.Append(levelText122);
            level122.Append(levelJustification122);
            level122.Append(previousParagraphProperties122);

            Level level123 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue123 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat123 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText123 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification123 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties123 = new PreviousParagraphProperties();
            Indentation indentation165 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties123.Append(indentation165);

            level123.Append(startNumberingValue123);
            level123.Append(numberingFormat123);
            level123.Append(levelText123);
            level123.Append(levelJustification123);
            level123.Append(previousParagraphProperties123);

            Level level124 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue124 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat124 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText124 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification124 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties124 = new PreviousParagraphProperties();
            Indentation indentation166 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties124.Append(indentation166);

            level124.Append(startNumberingValue124);
            level124.Append(numberingFormat124);
            level124.Append(levelText124);
            level124.Append(levelJustification124);
            level124.Append(previousParagraphProperties124);

            Level level125 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue125 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat125 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText125 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification125 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties125 = new PreviousParagraphProperties();
            Indentation indentation167 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties125.Append(indentation167);

            level125.Append(startNumberingValue125);
            level125.Append(numberingFormat125);
            level125.Append(levelText125);
            level125.Append(levelJustification125);
            level125.Append(previousParagraphProperties125);

            Level level126 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue126 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat126 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText126 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification126 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties126 = new PreviousParagraphProperties();
            Indentation indentation168 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties126.Append(indentation168);

            level126.Append(startNumberingValue126);
            level126.Append(numberingFormat126);
            level126.Append(levelText126);
            level126.Append(levelJustification126);
            level126.Append(previousParagraphProperties126);

            abstractNum14.Append(nsid14);
            abstractNum14.Append(multiLevelType14);
            abstractNum14.Append(templateCode14);
            abstractNum14.Append(level118);
            abstractNum14.Append(level119);
            abstractNum14.Append(level120);
            abstractNum14.Append(level121);
            abstractNum14.Append(level122);
            abstractNum14.Append(level123);
            abstractNum14.Append(level124);
            abstractNum14.Append(level125);
            abstractNum14.Append(level126);

            AbstractNum abstractNum15 = new AbstractNum() { AbstractNumberId = 14 };
            abstractNum15.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid15 = new Nsid() { Val = "72347E41" };
            MultiLevelType multiLevelType15 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode15 = new TemplateCode() { Val = "D58AC6F4" };

            Level level127 = new Level() { LevelIndex = 0, TemplateCode = "1046A65E" };
            StartNumberingValue startNumberingValue127 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat127 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText127 = new LevelText() { Val = "(%1)" };
            LevelJustification levelJustification127 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties127 = new PreviousParagraphProperties();

            Tabs tabs48 = new Tabs();
            TabStop tabStop51 = new TabStop() { Val = TabStopValues.Number, Position = 1284 };

            tabs48.Append(tabStop51);
            Indentation indentation169 = new Indentation() { Start = "1284", Hanging = "720" };

            previousParagraphProperties127.Append(tabs48);
            previousParagraphProperties127.Append(indentation169);

            NumberingSymbolRunProperties numberingSymbolRunProperties22 = new NumberingSymbolRunProperties();
            RunFonts runFonts429 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties22.Append(runFonts429);

            level127.Append(startNumberingValue127);
            level127.Append(numberingFormat127);
            level127.Append(levelText127);
            level127.Append(levelJustification127);
            level127.Append(previousParagraphProperties127);
            level127.Append(numberingSymbolRunProperties22);

            Level level128 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue128 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat128 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText128 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification128 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties128 = new PreviousParagraphProperties();

            Tabs tabs49 = new Tabs();
            TabStop tabStop52 = new TabStop() { Val = TabStopValues.Number, Position = 1524 };

            tabs49.Append(tabStop52);
            Indentation indentation170 = new Indentation() { Start = "1524", Hanging = "480" };

            previousParagraphProperties128.Append(tabs49);
            previousParagraphProperties128.Append(indentation170);

            level128.Append(startNumberingValue128);
            level128.Append(numberingFormat128);
            level128.Append(levelText128);
            level128.Append(levelJustification128);
            level128.Append(previousParagraphProperties128);

            Level level129 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue129 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat129 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText129 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification129 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties129 = new PreviousParagraphProperties();

            Tabs tabs50 = new Tabs();
            TabStop tabStop53 = new TabStop() { Val = TabStopValues.Number, Position = 2004 };

            tabs50.Append(tabStop53);
            Indentation indentation171 = new Indentation() { Start = "2004", Hanging = "480" };

            previousParagraphProperties129.Append(tabs50);
            previousParagraphProperties129.Append(indentation171);

            level129.Append(startNumberingValue129);
            level129.Append(numberingFormat129);
            level129.Append(levelText129);
            level129.Append(levelJustification129);
            level129.Append(previousParagraphProperties129);

            Level level130 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue130 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat130 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText130 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification130 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties130 = new PreviousParagraphProperties();

            Tabs tabs51 = new Tabs();
            TabStop tabStop54 = new TabStop() { Val = TabStopValues.Number, Position = 2484 };

            tabs51.Append(tabStop54);
            Indentation indentation172 = new Indentation() { Start = "2484", Hanging = "480" };

            previousParagraphProperties130.Append(tabs51);
            previousParagraphProperties130.Append(indentation172);

            level130.Append(startNumberingValue130);
            level130.Append(numberingFormat130);
            level130.Append(levelText130);
            level130.Append(levelJustification130);
            level130.Append(previousParagraphProperties130);

            Level level131 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue131 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat131 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText131 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification131 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties131 = new PreviousParagraphProperties();

            Tabs tabs52 = new Tabs();
            TabStop tabStop55 = new TabStop() { Val = TabStopValues.Number, Position = 2964 };

            tabs52.Append(tabStop55);
            Indentation indentation173 = new Indentation() { Start = "2964", Hanging = "480" };

            previousParagraphProperties131.Append(tabs52);
            previousParagraphProperties131.Append(indentation173);

            level131.Append(startNumberingValue131);
            level131.Append(numberingFormat131);
            level131.Append(levelText131);
            level131.Append(levelJustification131);
            level131.Append(previousParagraphProperties131);

            Level level132 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue132 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat132 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText132 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification132 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties132 = new PreviousParagraphProperties();

            Tabs tabs53 = new Tabs();
            TabStop tabStop56 = new TabStop() { Val = TabStopValues.Number, Position = 3444 };

            tabs53.Append(tabStop56);
            Indentation indentation174 = new Indentation() { Start = "3444", Hanging = "480" };

            previousParagraphProperties132.Append(tabs53);
            previousParagraphProperties132.Append(indentation174);

            level132.Append(startNumberingValue132);
            level132.Append(numberingFormat132);
            level132.Append(levelText132);
            level132.Append(levelJustification132);
            level132.Append(previousParagraphProperties132);

            Level level133 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue133 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat133 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText133 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification133 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties133 = new PreviousParagraphProperties();

            Tabs tabs54 = new Tabs();
            TabStop tabStop57 = new TabStop() { Val = TabStopValues.Number, Position = 3924 };

            tabs54.Append(tabStop57);
            Indentation indentation175 = new Indentation() { Start = "3924", Hanging = "480" };

            previousParagraphProperties133.Append(tabs54);
            previousParagraphProperties133.Append(indentation175);

            level133.Append(startNumberingValue133);
            level133.Append(numberingFormat133);
            level133.Append(levelText133);
            level133.Append(levelJustification133);
            level133.Append(previousParagraphProperties133);

            Level level134 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue134 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat134 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText134 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification134 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties134 = new PreviousParagraphProperties();

            Tabs tabs55 = new Tabs();
            TabStop tabStop58 = new TabStop() { Val = TabStopValues.Number, Position = 4404 };

            tabs55.Append(tabStop58);
            Indentation indentation176 = new Indentation() { Start = "4404", Hanging = "480" };

            previousParagraphProperties134.Append(tabs55);
            previousParagraphProperties134.Append(indentation176);

            level134.Append(startNumberingValue134);
            level134.Append(numberingFormat134);
            level134.Append(levelText134);
            level134.Append(levelJustification134);
            level134.Append(previousParagraphProperties134);

            Level level135 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue135 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat135 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText135 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification135 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties135 = new PreviousParagraphProperties();

            Tabs tabs56 = new Tabs();
            TabStop tabStop59 = new TabStop() { Val = TabStopValues.Number, Position = 4884 };

            tabs56.Append(tabStop59);
            Indentation indentation177 = new Indentation() { Start = "4884", Hanging = "480" };

            previousParagraphProperties135.Append(tabs56);
            previousParagraphProperties135.Append(indentation177);

            level135.Append(startNumberingValue135);
            level135.Append(numberingFormat135);
            level135.Append(levelText135);
            level135.Append(levelJustification135);
            level135.Append(previousParagraphProperties135);

            abstractNum15.Append(nsid15);
            abstractNum15.Append(multiLevelType15);
            abstractNum15.Append(templateCode15);
            abstractNum15.Append(level127);
            abstractNum15.Append(level128);
            abstractNum15.Append(level129);
            abstractNum15.Append(level130);
            abstractNum15.Append(level131);
            abstractNum15.Append(level132);
            abstractNum15.Append(level133);
            abstractNum15.Append(level134);
            abstractNum15.Append(level135);

            AbstractNum abstractNum16 = new AbstractNum() { AbstractNumberId = 15 };
            abstractNum16.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid16 = new Nsid() { Val = "72F72D62" };
            MultiLevelType multiLevelType16 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode16 = new TemplateCode() { Val = "C03079DE" };

            Level level136 = new Level() { LevelIndex = 0, TemplateCode = "F4EEFD4A" };
            StartNumberingValue startNumberingValue136 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat136 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText136 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification136 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties136 = new PreviousParagraphProperties();
            Indentation indentation178 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties136.Append(indentation178);

            NumberingSymbolRunProperties numberingSymbolRunProperties23 = new NumberingSymbolRunProperties();
            RunFonts runFonts430 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties23.Append(runFonts430);

            level136.Append(startNumberingValue136);
            level136.Append(numberingFormat136);
            level136.Append(levelText136);
            level136.Append(levelJustification136);
            level136.Append(previousParagraphProperties136);
            level136.Append(numberingSymbolRunProperties23);

            Level level137 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue137 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat137 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText137 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification137 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties137 = new PreviousParagraphProperties();
            Indentation indentation179 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties137.Append(indentation179);

            level137.Append(startNumberingValue137);
            level137.Append(numberingFormat137);
            level137.Append(levelText137);
            level137.Append(levelJustification137);
            level137.Append(previousParagraphProperties137);

            Level level138 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue138 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat138 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText138 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification138 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties138 = new PreviousParagraphProperties();
            Indentation indentation180 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties138.Append(indentation180);

            level138.Append(startNumberingValue138);
            level138.Append(numberingFormat138);
            level138.Append(levelText138);
            level138.Append(levelJustification138);
            level138.Append(previousParagraphProperties138);

            Level level139 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue139 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat139 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText139 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification139 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties139 = new PreviousParagraphProperties();
            Indentation indentation181 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties139.Append(indentation181);

            level139.Append(startNumberingValue139);
            level139.Append(numberingFormat139);
            level139.Append(levelText139);
            level139.Append(levelJustification139);
            level139.Append(previousParagraphProperties139);

            Level level140 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue140 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat140 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText140 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification140 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties140 = new PreviousParagraphProperties();
            Indentation indentation182 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties140.Append(indentation182);

            level140.Append(startNumberingValue140);
            level140.Append(numberingFormat140);
            level140.Append(levelText140);
            level140.Append(levelJustification140);
            level140.Append(previousParagraphProperties140);

            Level level141 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue141 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat141 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText141 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification141 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties141 = new PreviousParagraphProperties();
            Indentation indentation183 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties141.Append(indentation183);

            level141.Append(startNumberingValue141);
            level141.Append(numberingFormat141);
            level141.Append(levelText141);
            level141.Append(levelJustification141);
            level141.Append(previousParagraphProperties141);

            Level level142 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue142 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat142 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText142 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification142 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties142 = new PreviousParagraphProperties();
            Indentation indentation184 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties142.Append(indentation184);

            level142.Append(startNumberingValue142);
            level142.Append(numberingFormat142);
            level142.Append(levelText142);
            level142.Append(levelJustification142);
            level142.Append(previousParagraphProperties142);

            Level level143 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue143 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat143 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText143 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification143 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties143 = new PreviousParagraphProperties();
            Indentation indentation185 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties143.Append(indentation185);

            level143.Append(startNumberingValue143);
            level143.Append(numberingFormat143);
            level143.Append(levelText143);
            level143.Append(levelJustification143);
            level143.Append(previousParagraphProperties143);

            Level level144 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue144 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat144 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText144 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification144 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties144 = new PreviousParagraphProperties();
            Indentation indentation186 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties144.Append(indentation186);

            level144.Append(startNumberingValue144);
            level144.Append(numberingFormat144);
            level144.Append(levelText144);
            level144.Append(levelJustification144);
            level144.Append(previousParagraphProperties144);

            abstractNum16.Append(nsid16);
            abstractNum16.Append(multiLevelType16);
            abstractNum16.Append(templateCode16);
            abstractNum16.Append(level136);
            abstractNum16.Append(level137);
            abstractNum16.Append(level138);
            abstractNum16.Append(level139);
            abstractNum16.Append(level140);
            abstractNum16.Append(level141);
            abstractNum16.Append(level142);
            abstractNum16.Append(level143);
            abstractNum16.Append(level144);

            AbstractNum abstractNum17 = new AbstractNum() { AbstractNumberId = 16 };
            abstractNum17.SetAttribute(new OpenXmlAttribute("w15", "restartNumberingAfterBreak", "http://schemas.microsoft.com/office/word/2012/wordml", "0"));
            Nsid nsid17 = new Nsid() { Val = "7BAF4830" };
            MultiLevelType multiLevelType17 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode17 = new TemplateCode() { Val = "4C9690DC" };

            Level level145 = new Level() { LevelIndex = 0, TemplateCode = "6EF420B6" };
            StartNumberingValue startNumberingValue145 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat145 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText145 = new LevelText() { Val = "Ÿ" };
            LevelJustification levelJustification145 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties145 = new PreviousParagraphProperties();
            Indentation indentation187 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties145.Append(indentation187);

            NumberingSymbolRunProperties numberingSymbolRunProperties24 = new NumberingSymbolRunProperties();
            RunFonts runFonts431 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties24.Append(runFonts431);

            level145.Append(startNumberingValue145);
            level145.Append(numberingFormat145);
            level145.Append(levelText145);
            level145.Append(levelJustification145);
            level145.Append(previousParagraphProperties145);
            level145.Append(numberingSymbolRunProperties24);

            Level level146 = new Level() { LevelIndex = 1, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue146 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat146 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText146 = new LevelText() { Val = "n" };
            LevelJustification levelJustification146 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties146 = new PreviousParagraphProperties();
            Indentation indentation188 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties146.Append(indentation188);

            NumberingSymbolRunProperties numberingSymbolRunProperties25 = new NumberingSymbolRunProperties();
            RunFonts runFonts432 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties25.Append(runFonts432);

            level146.Append(startNumberingValue146);
            level146.Append(numberingFormat146);
            level146.Append(levelText146);
            level146.Append(levelJustification146);
            level146.Append(previousParagraphProperties146);
            level146.Append(numberingSymbolRunProperties25);

            Level level147 = new Level() { LevelIndex = 2, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue147 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat147 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText147 = new LevelText() { Val = "u" };
            LevelJustification levelJustification147 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties147 = new PreviousParagraphProperties();
            Indentation indentation189 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties147.Append(indentation189);

            NumberingSymbolRunProperties numberingSymbolRunProperties26 = new NumberingSymbolRunProperties();
            RunFonts runFonts433 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties26.Append(runFonts433);

            level147.Append(startNumberingValue147);
            level147.Append(numberingFormat147);
            level147.Append(levelText147);
            level147.Append(levelJustification147);
            level147.Append(previousParagraphProperties147);
            level147.Append(numberingSymbolRunProperties26);

            Level level148 = new Level() { LevelIndex = 3, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue148 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat148 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText148 = new LevelText() { Val = "l" };
            LevelJustification levelJustification148 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties148 = new PreviousParagraphProperties();
            Indentation indentation190 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties148.Append(indentation190);

            NumberingSymbolRunProperties numberingSymbolRunProperties27 = new NumberingSymbolRunProperties();
            RunFonts runFonts434 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties27.Append(runFonts434);

            level148.Append(startNumberingValue148);
            level148.Append(numberingFormat148);
            level148.Append(levelText148);
            level148.Append(levelJustification148);
            level148.Append(previousParagraphProperties148);
            level148.Append(numberingSymbolRunProperties27);

            Level level149 = new Level() { LevelIndex = 4, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue149 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat149 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText149 = new LevelText() { Val = "n" };
            LevelJustification levelJustification149 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties149 = new PreviousParagraphProperties();
            Indentation indentation191 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties149.Append(indentation191);

            NumberingSymbolRunProperties numberingSymbolRunProperties28 = new NumberingSymbolRunProperties();
            RunFonts runFonts435 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties28.Append(runFonts435);

            level149.Append(startNumberingValue149);
            level149.Append(numberingFormat149);
            level149.Append(levelText149);
            level149.Append(levelJustification149);
            level149.Append(previousParagraphProperties149);
            level149.Append(numberingSymbolRunProperties28);

            Level level150 = new Level() { LevelIndex = 5, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue150 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat150 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText150 = new LevelText() { Val = "u" };
            LevelJustification levelJustification150 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties150 = new PreviousParagraphProperties();
            Indentation indentation192 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties150.Append(indentation192);

            NumberingSymbolRunProperties numberingSymbolRunProperties29 = new NumberingSymbolRunProperties();
            RunFonts runFonts436 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties29.Append(runFonts436);

            level150.Append(startNumberingValue150);
            level150.Append(numberingFormat150);
            level150.Append(levelText150);
            level150.Append(levelJustification150);
            level150.Append(previousParagraphProperties150);
            level150.Append(numberingSymbolRunProperties29);

            Level level151 = new Level() { LevelIndex = 6, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue151 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat151 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText151 = new LevelText() { Val = "l" };
            LevelJustification levelJustification151 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties151 = new PreviousParagraphProperties();
            Indentation indentation193 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties151.Append(indentation193);

            NumberingSymbolRunProperties numberingSymbolRunProperties30 = new NumberingSymbolRunProperties();
            RunFonts runFonts437 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties30.Append(runFonts437);

            level151.Append(startNumberingValue151);
            level151.Append(numberingFormat151);
            level151.Append(levelText151);
            level151.Append(levelJustification151);
            level151.Append(previousParagraphProperties151);
            level151.Append(numberingSymbolRunProperties30);

            Level level152 = new Level() { LevelIndex = 7, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue152 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat152 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText152 = new LevelText() { Val = "n" };
            LevelJustification levelJustification152 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties152 = new PreviousParagraphProperties();
            Indentation indentation194 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties152.Append(indentation194);

            NumberingSymbolRunProperties numberingSymbolRunProperties31 = new NumberingSymbolRunProperties();
            RunFonts runFonts438 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties31.Append(runFonts438);

            level152.Append(startNumberingValue152);
            level152.Append(numberingFormat152);
            level152.Append(levelText152);
            level152.Append(levelJustification152);
            level152.Append(previousParagraphProperties152);
            level152.Append(numberingSymbolRunProperties31);

            Level level153 = new Level() { LevelIndex = 8, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue153 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat153 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText153 = new LevelText() { Val = "u" };
            LevelJustification levelJustification153 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties153 = new PreviousParagraphProperties();
            Indentation indentation195 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties153.Append(indentation195);

            NumberingSymbolRunProperties numberingSymbolRunProperties32 = new NumberingSymbolRunProperties();
            RunFonts runFonts439 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties32.Append(runFonts439);

            level153.Append(startNumberingValue153);
            level153.Append(numberingFormat153);
            level153.Append(levelText153);
            level153.Append(levelJustification153);
            level153.Append(previousParagraphProperties153);
            level153.Append(numberingSymbolRunProperties32);

            abstractNum17.Append(nsid17);
            abstractNum17.Append(multiLevelType17);
            abstractNum17.Append(templateCode17);
            abstractNum17.Append(level145);
            abstractNum17.Append(level146);
            abstractNum17.Append(level147);
            abstractNum17.Append(level148);
            abstractNum17.Append(level149);
            abstractNum17.Append(level150);
            abstractNum17.Append(level151);
            abstractNum17.Append(level152);
            abstractNum17.Append(level153);

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            numberingInstance1.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "595329676"));
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 7 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            numberingInstance2.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1812214147"));
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 12 };

            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            numberingInstance3.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "504057109"));
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 6 };

            numberingInstance3.Append(abstractNumId3);

            NumberingInstance numberingInstance4 = new NumberingInstance() { NumberID = 4 };
            numberingInstance4.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "408386300"));
            AbstractNumId abstractNumId4 = new AbstractNumId() { Val = 2 };

            numberingInstance4.Append(abstractNumId4);

            NumberingInstance numberingInstance5 = new NumberingInstance() { NumberID = 5 };
            numberingInstance5.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "47150400"));
            AbstractNumId abstractNumId5 = new AbstractNumId() { Val = 10 };

            numberingInstance5.Append(abstractNumId5);

            NumberingInstance numberingInstance6 = new NumberingInstance() { NumberID = 6 };
            numberingInstance6.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "2138257118"));
            AbstractNumId abstractNumId6 = new AbstractNumId() { Val = 15 };

            numberingInstance6.Append(abstractNumId6);

            NumberingInstance numberingInstance7 = new NumberingInstance() { NumberID = 7 };
            numberingInstance7.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "597255985"));
            AbstractNumId abstractNumId7 = new AbstractNumId() { Val = 16 };

            numberingInstance7.Append(abstractNumId7);

            NumberingInstance numberingInstance8 = new NumberingInstance() { NumberID = 8 };
            numberingInstance8.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "100151801"));
            AbstractNumId abstractNumId8 = new AbstractNumId() { Val = 5 };

            numberingInstance8.Append(abstractNumId8);

            NumberingInstance numberingInstance9 = new NumberingInstance() { NumberID = 9 };
            numberingInstance9.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1114133431"));
            AbstractNumId abstractNumId9 = new AbstractNumId() { Val = 14 };

            numberingInstance9.Append(abstractNumId9);

            NumberingInstance numberingInstance10 = new NumberingInstance() { NumberID = 10 };
            numberingInstance10.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "297876902"));
            AbstractNumId abstractNumId10 = new AbstractNumId() { Val = 13 };

            numberingInstance10.Append(abstractNumId10);

            NumberingInstance numberingInstance11 = new NumberingInstance() { NumberID = 11 };
            numberingInstance11.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "241723705"));
            AbstractNumId abstractNumId11 = new AbstractNumId() { Val = 1 };

            numberingInstance11.Append(abstractNumId11);

            NumberingInstance numberingInstance12 = new NumberingInstance() { NumberID = 12 };
            numberingInstance12.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1707099846"));
            AbstractNumId abstractNumId12 = new AbstractNumId() { Val = 4 };

            numberingInstance12.Append(abstractNumId12);

            NumberingInstance numberingInstance13 = new NumberingInstance() { NumberID = 13 };
            numberingInstance13.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1224411702"));
            AbstractNumId abstractNumId13 = new AbstractNumId() { Val = 9 };

            numberingInstance13.Append(abstractNumId13);

            NumberingInstance numberingInstance14 = new NumberingInstance() { NumberID = 14 };
            numberingInstance14.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1523741738"));
            AbstractNumId abstractNumId14 = new AbstractNumId() { Val = 8 };

            numberingInstance14.Append(abstractNumId14);

            NumberingInstance numberingInstance15 = new NumberingInstance() { NumberID = 15 };
            numberingInstance15.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "501243550"));
            AbstractNumId abstractNumId15 = new AbstractNumId() { Val = 3 };

            numberingInstance15.Append(abstractNumId15);

            NumberingInstance numberingInstance16 = new NumberingInstance() { NumberID = 16 };
            numberingInstance16.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "413165727"));
            AbstractNumId abstractNumId16 = new AbstractNumId() { Val = 11 };

            numberingInstance16.Append(abstractNumId16);

            NumberingInstance numberingInstance17 = new NumberingInstance() { NumberID = 17 };
            numberingInstance17.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1756659641"));
            AbstractNumId abstractNumId17 = new AbstractNumId() { Val = 0 };

            numberingInstance17.Append(abstractNumId17);

            numbering1.Append(abstractNum1);
            numbering1.Append(abstractNum2);
            numbering1.Append(abstractNum3);
            numbering1.Append(abstractNum4);
            numbering1.Append(abstractNum5);
            numbering1.Append(abstractNum6);
            numbering1.Append(abstractNum7);
            numbering1.Append(abstractNum8);
            numbering1.Append(abstractNum9);
            numbering1.Append(abstractNum10);
            numbering1.Append(abstractNum11);
            numbering1.Append(abstractNum12);
            numbering1.Append(abstractNum13);
            numbering1.Append(abstractNum14);
            numbering1.Append(abstractNum15);
            numbering1.Append(abstractNum16);
            numbering1.Append(abstractNum17);
            numbering1.Append(numberingInstance1);
            numbering1.Append(numberingInstance2);
            numbering1.Append(numberingInstance3);
            numbering1.Append(numberingInstance4);
            numbering1.Append(numberingInstance5);
            numbering1.Append(numberingInstance6);
            numbering1.Append(numberingInstance7);
            numbering1.Append(numberingInstance8);
            numbering1.Append(numberingInstance9);
            numbering1.Append(numberingInstance10);
            numbering1.Append(numberingInstance11);
            numbering1.Append(numberingInstance12);
            numbering1.Append(numberingInstance13);
            numbering1.Append(numberingInstance14);
            numbering1.Append(numberingInstance15);
            numbering1.Append(numberingInstance16);
            numbering1.Append(numberingInstance17);

            numberingDefinitionsPart1.Numbering = numbering1;
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
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

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
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" };
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
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
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
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline2 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill2);
            outline2.Append(presetDash1);
            outline2.Append(miter1);

            A.Outline outline3 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill3);
            outline3.Append(presetDash2);
            outline3.Append(miter2);

            A.Outline outline4 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline4.Append(solidFill4);
            outline4.Append(presetDash3);
            outline4.Append(miter3);

            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
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

            Paragraph paragraph86 = new Paragraph() { RsidParagraphAddition = "00485F68", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00485F68", ParagraphId = "30A97FF2", TextId = "77777777" };

            Run run333 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run333.Append(separatorMark1);

            paragraph86.Append(run333);

            endnote1.Append(paragraph86);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph87 = new Paragraph() { RsidParagraphAddition = "00485F68", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00485F68", ParagraphId = "5F29DDA7", TextId = "77777777" };

            Run run334 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run334.Append(continuationSeparatorMark1);

            paragraph87.Append(run334);

            endnote2.Append(paragraph87);

            endnotes1.Append(endnote1);
            endnotes1.Append(endnote2);

            endnotesPart1.Endnotes = endnotes1;
        }

        // Generates content of customXmlPart4.
        private void GenerateCustomXmlPart4Content(CustomXmlPart customXmlPart4)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart4.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"no\"?><b:Sources SelectedStyle=\"\\APA.XSL\" StyleName=\"APA\" xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\" xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\"></b:Sources>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart4.
        private void GenerateCustomXmlPropertiesPart4Content(CustomXmlPropertiesPart customXmlPropertiesPart4)
        {
            Ds.DataStoreItem dataStoreItem4 = new Ds.DataStoreItem() { ItemId = "{BED0393C-0D48-4566-9762-26CCCC325536}" };
            dataStoreItem4.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences4 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference18 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography" };

            schemaReferences4.Append(schemaReference18);

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

            Paragraph paragraph88 = new Paragraph() { RsidParagraphAddition = "00485F68", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00485F68", ParagraphId = "76AAAEAF", TextId = "77777777" };

            Run run335 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run335.Append(separatorMark2);

            paragraph88.Append(run335);

            footnote1.Append(paragraph88);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph89 = new Paragraph() { RsidParagraphAddition = "00485F68", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00485F68", ParagraphId = "4939AF8F", TextId = "77777777" };

            Run run336 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run336.Append(continuationSeparatorMark2);

            paragraph89.Append(run336);

            footnote2.Append(paragraph89);

            Footnote footnote3 = new Footnote() { Id = 1 };

            Paragraph paragraph90 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00D43514", RsidParagraphProperties = "001C3E6D", RsidRunAdditionDefault = "006C3779", ParagraphId = "0D28FF3D", TextId = "77777777" };

            ParagraphProperties paragraphProperties86 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation196 = new Indentation() { Start = "64", Hanging = "64", HangingChars = 40 };

            ParagraphMarkRunProperties paragraphMarkRunProperties82 = new ParagraphMarkRunProperties();
            FontSize fontSize354 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties82.Append(fontSize354);
            paragraphMarkRunProperties82.Append(fontSizeComplexScript90);

            paragraphProperties86.Append(spacingBetweenLines23);
            paragraphProperties86.Append(indentation196);
            paragraphProperties86.Append(paragraphMarkRunProperties82);

            Run run337 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties329 = new RunProperties();
            RunStyle runStyle29 = new RunStyle() { Val = "a9" };
            FontSize fontSize355 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "16" };

            runProperties329.Append(runStyle29);
            runProperties329.Append(fontSize355);
            runProperties329.Append(fontSizeComplexScript91);
            FootnoteReferenceMark footnoteReferenceMark1 = new FootnoteReferenceMark();

            run337.Append(runProperties329);
            run337.Append(footnoteReferenceMark1);

            Run run338 = new Run() { RsidRunAddition = "000D4BC6" };

            RunProperties runProperties330 = new RunProperties();
            FontSize fontSize356 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "16" };

            runProperties330.Append(fontSize356);
            runProperties330.Append(fontSizeComplexScript92);
            Text text191 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text191.Text = " List the key regulatory changes from the end of the previous internal audit until the current audit cut-off date to inform the auditor of any siginificant changes that have occurred";

            run338.Append(runProperties330);
            run338.Append(text191);

            Run run339 = new Run() { RsidRunAddition = "000D4BC6" };

            RunProperties runProperties331 = new RunProperties();
            RunFonts runFonts440 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize357 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "16" };

            runProperties331.Append(runFonts440);
            runProperties331.Append(fontSize357);
            runProperties331.Append(fontSizeComplexScript93);
            Text text192 = new Text();
            text192.Text = ".";

            run339.Append(runProperties331);
            run339.Append(text192);

            paragraph90.Append(paragraphProperties86);
            paragraph90.Append(run337);
            paragraph90.Append(run338);
            paragraph90.Append(run339);

            footnote3.Append(paragraph90);

            Footnote footnote4 = new Footnote() { Id = 2 };

            Paragraph paragraph91 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00CF43E8", RsidRunAdditionDefault = "00CF43E8", ParagraphId = "78D47696", TextId = "77777777" };

            ParagraphProperties paragraphProperties87 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "a7" };

            ParagraphMarkRunProperties paragraphMarkRunProperties83 = new ParagraphMarkRunProperties();
            FontSize fontSize358 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties83.Append(fontSize358);
            paragraphMarkRunProperties83.Append(fontSizeComplexScript94);

            paragraphProperties87.Append(paragraphStyleId9);
            paragraphProperties87.Append(paragraphMarkRunProperties83);

            Run run340 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties332 = new RunProperties();
            RunStyle runStyle30 = new RunStyle() { Val = "a9" };
            FontSize fontSize359 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "16" };

            runProperties332.Append(runStyle30);
            runProperties332.Append(fontSize359);
            runProperties332.Append(fontSizeComplexScript95);
            FootnoteReferenceMark footnoteReferenceMark2 = new FootnoteReferenceMark();

            run340.Append(runProperties332);
            run340.Append(footnoteReferenceMark2);

            Run run341 = new Run() { RsidRunAddition = "004E428D" };

            RunProperties runProperties333 = new RunProperties();
            RunFonts runFonts441 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize360 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "16" };

            runProperties333.Append(runFonts441);
            runProperties333.Append(fontSize360);
            runProperties333.Append(fontSizeComplexScript96);
            Text text193 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text193.Text = " If there are no major issues identified in previous audit";

            run341.Append(runProperties333);
            run341.Append(text193);

            Run run342 = new Run() { RsidRunAddition = "00534826" };

            RunProperties runProperties334 = new RunProperties();
            FontSize fontSize361 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "16" };

            runProperties334.Append(fontSize361);
            runProperties334.Append(fontSizeComplexScript97);
            Text text194 = new Text();
            text194.Text = "s";

            run342.Append(runProperties334);
            run342.Append(text194);

            Run run343 = new Run() { RsidRunAddition = "00534826" };

            RunProperties runProperties335 = new RunProperties();
            RunFonts runFonts442 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize362 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "16" };

            runProperties335.Append(runFonts442);
            runProperties335.Append(fontSize362);
            runProperties335.Append(fontSizeComplexScript98);
            Text text195 = new Text();
            text195.Text = "/";

            run343.Append(runProperties335);
            run343.Append(text195);

            Run run344 = new Run() { RsidRunAddition = "00534826" };

            RunProperties runProperties336 = new RunProperties();
            FontSize fontSize363 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "16" };

            runProperties336.Append(fontSize363);
            runProperties336.Append(fontSizeComplexScript99);
            Text text196 = new Text();
            text196.Text = "examinations";

            run344.Append(runProperties336);
            run344.Append(text196);

            Run run345 = new Run() { RsidRunAddition = "004E428D" };

            RunProperties runProperties337 = new RunProperties();
            FontSize fontSize364 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript100 = new FontSizeComplexScript() { Val = "16" };

            runProperties337.Append(fontSize364);
            runProperties337.Append(fontSizeComplexScript100);
            Text text197 = new Text();
            text197.Text = ", please mark as “";

            run345.Append(runProperties337);
            run345.Append(text197);

            Run run346 = new Run() { RsidRunAddition = "004E428D" };

            RunProperties runProperties338 = new RunProperties();
            RunFonts runFonts443 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize365 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript101 = new FontSizeComplexScript() { Val = "16" };

            runProperties338.Append(runFonts443);
            runProperties338.Append(fontSize365);
            runProperties338.Append(fontSizeComplexScript101);
            Text text198 = new Text();
            text198.Text = "N/A";

            run346.Append(runProperties338);
            run346.Append(text198);

            Run run347 = new Run() { RsidRunAddition = "004E428D" };

            RunProperties runProperties339 = new RunProperties();
            FontSize fontSize366 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "16" };

            runProperties339.Append(fontSize366);
            runProperties339.Append(fontSizeComplexScript102);
            Text text199 = new Text();
            text199.Text = "”";

            run347.Append(runProperties339);
            run347.Append(text199);

            Run run348 = new Run() { RsidRunAddition = "004E428D" };

            RunProperties runProperties340 = new RunProperties();
            RunFonts runFonts444 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize367 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "16" };

            runProperties340.Append(runFonts444);
            runProperties340.Append(fontSize367);
            runProperties340.Append(fontSizeComplexScript103);
            Text text200 = new Text();
            text200.Text = ".";

            run348.Append(runProperties340);
            run348.Append(text200);

            paragraph91.Append(paragraphProperties87);
            paragraph91.Append(run340);
            paragraph91.Append(run341);
            paragraph91.Append(run342);
            paragraph91.Append(run343);
            paragraph91.Append(run344);
            paragraph91.Append(run345);
            paragraph91.Append(run346);
            paragraph91.Append(run347);
            paragraph91.Append(run348);

            footnote4.Append(paragraph91);

            Footnote footnote5 = new Footnote() { Id = 3 };

            Paragraph paragraph92 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F03C81", RsidParagraphProperties = "00F03C81", RsidRunAdditionDefault = "00F03C81", ParagraphId = "09D657E0", TextId = "328C4052" };

            ParagraphProperties paragraphProperties88 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines24 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation197 = new Indentation() { Start = "158", Hanging = "158", HangingChars = 99 };

            ParagraphMarkRunProperties paragraphMarkRunProperties84 = new ParagraphMarkRunProperties();
            FontSize fontSize368 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties84.Append(fontSize368);
            paragraphMarkRunProperties84.Append(fontSizeComplexScript104);

            paragraphProperties88.Append(spacingBetweenLines24);
            paragraphProperties88.Append(indentation197);
            paragraphProperties88.Append(paragraphMarkRunProperties84);

            Run run349 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties341 = new RunProperties();
            RunStyle runStyle31 = new RunStyle() { Val = "a9" };
            FontSize fontSize369 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "16" };

            runProperties341.Append(runStyle31);
            runProperties341.Append(fontSize369);
            runProperties341.Append(fontSizeComplexScript105);
            FootnoteReferenceMark footnoteReferenceMark3 = new FootnoteReferenceMark();

            run349.Append(runProperties341);
            run349.Append(footnoteReferenceMark3);

            Run run350 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties342 = new RunProperties();
            FontSize fontSize370 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript106 = new FontSizeComplexScript() { Val = "16" };

            runProperties342.Append(fontSize370);
            runProperties342.Append(fontSizeComplexScript106);
            Text text201 = new Text();
            text201.Text = "a)";

            run350.Append(runProperties342);
            run350.Append(text201);

            Run run351 = new Run() { RsidRunProperties = "006D564F", RsidRunAddition = "006D564F" };

            RunProperties runProperties343 = new RunProperties();
            RunFonts runFonts445 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize371 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript107 = new FontSizeComplexScript() { Val = "16" };

            runProperties343.Append(runFonts445);
            runProperties343.Append(fontSize371);
            runProperties343.Append(fontSizeComplexScript107);
            Text text202 = new Text();
            text202.Text = "T";

            run351.Append(runProperties343);
            run351.Append(text202);

            Run run352 = new Run() { RsidRunProperties = "006D564F", RsidRunAddition = "006D564F" };

            RunProperties runProperties344 = new RunProperties();
            FontSize fontSize372 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "16" };

            runProperties344.Append(fontSize372);
            runProperties344.Append(fontSizeComplexScript108);
            Text text203 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text203.Text = "he compliance of ";

            run352.Append(runProperties344);
            run352.Append(text203);

            Run run353 = new Run() { RsidRunProperties = "006D564F", RsidRunAddition = "00153D64" };

            RunProperties runProperties345 = new RunProperties();
            FontSize fontSize373 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "16" };

            runProperties345.Append(fontSize373);
            runProperties345.Append(fontSizeComplexScript109);
            Text text204 = new Text();
            text204.Text = "competent";

            run353.Append(runProperties345);
            run353.Append(text204);

            Run run354 = new Run() { RsidRunProperties = "006D564F", RsidRunAddition = "006D564F" };

            RunProperties runProperties346 = new RunProperties();
            FontSize fontSize374 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "16" };

            runProperties346.Append(fontSize374);
            runProperties346.Append(fontSizeComplexScript110);
            Text text205 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text205.Text = " authority";

            run354.Append(runProperties346);
            run354.Append(text205);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run355 = new Run() { RsidRunProperties = "006D564F", RsidRunAddition = "006D564F" };

            RunProperties runProperties347 = new RunProperties();
            FontSize fontSize375 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "16" };

            runProperties347.Append(fontSize375);
            runProperties347.Append(fontSizeComplexScript111);
            Text text206 = new Text();
            text206.Text = "’";

            run355.Append(runProperties347);
            run355.Append(text206);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run356 = new Run() { RsidRunProperties = "006D564F", RsidRunAddition = "006D564F" };

            RunProperties runProperties348 = new RunProperties();
            FontSize fontSize376 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "16" };

            runProperties348.Append(fontSize376);
            runProperties348.Append(fontSizeComplexScript112);
            Text text207 = new Text();
            text207.Text = "s directives";

            run356.Append(runProperties348);
            run356.Append(text207);

            Run run357 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties349 = new RunProperties();
            FontSize fontSize377 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "16" };

            runProperties349.Append(fontSize377);
            runProperties349.Append(fontSizeComplexScript113);
            Text text208 = new Text();
            text208.Text = "：";

            run357.Append(runProperties349);
            run357.Append(text208);

            Run run358 = new Run() { RsidRunAddition = "00760421" };

            RunProperties runProperties350 = new RunProperties();
            RunFonts runFonts446 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize378 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "16" };

            runProperties350.Append(runFonts446);
            runProperties350.Append(fontSize378);
            runProperties350.Append(fontSizeComplexScript114);
            Text text209 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text209.Text = "this includes the commitments made by the auditee to the authority when applying for business or responding to ";

            run358.Append(runProperties350);
            run358.Append(text209);

            Run run359 = new Run() { RsidRunProperties = "00101E7E", RsidRunAddition = "00760421" };

            RunProperties runProperties351 = new RunProperties();
            RunFonts runFonts447 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize379 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript115 = new FontSizeComplexScript() { Val = "16" };

            runProperties351.Append(runFonts447);
            runProperties351.Append(fontSize379);
            runProperties351.Append(fontSizeComplexScript115);
            Text text210 = new Text();
            text210.Text = "impro";

            run359.Append(runProperties351);
            run359.Append(text210);

            Run run360 = new Run() { RsidRunProperties = "00101E7E", RsidRunAddition = "004C41E7" };

            RunProperties runProperties352 = new RunProperties();
            FontSize fontSize380 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript116 = new FontSizeComplexScript() { Val = "16" };

            runProperties352.Append(fontSize380);
            runProperties352.Append(fontSizeComplexScript116);
            Text text211 = new Text();
            text211.Text = "ve";

            run360.Append(runProperties352);
            run360.Append(text211);

            Run run361 = new Run() { RsidRunProperties = "00101E7E", RsidRunAddition = "00760421" };

            RunProperties runProperties353 = new RunProperties();
            RunFonts runFonts448 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize381 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript117 = new FontSizeComplexScript() { Val = "16" };

            runProperties353.Append(runFonts448);
            runProperties353.Append(fontSize381);
            runProperties353.Append(fontSizeComplexScript117);
            Text text212 = new Text();
            text212.Text = "ment progress";

            run361.Append(runProperties353);
            run361.Append(text212);

            Run run362 = new Run() { RsidRunAddition = "00760421" };

            RunProperties runProperties354 = new RunProperties();
            RunFonts runFonts449 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize382 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript118 = new FontSizeComplexScript() { Val = "16" };

            runProperties354.Append(runFonts449);
            runProperties354.Append(fontSize382);
            runProperties354.Append(fontSizeComplexScript118);
            Text text213 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text213.Text = ", ";

            run362.Append(runProperties354);
            run362.Append(text213);

            Run run363 = new Run() { RsidRunAddition = "00746DA3" };

            RunProperties runProperties355 = new RunProperties();
            FontSize fontSize383 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript119 = new FontSizeComplexScript() { Val = "16" };

            runProperties355.Append(fontSize383);
            runProperties355.Append(fontSizeComplexScript119);
            Text text214 = new Text();
            text214.Text = "as well as the authority’s written instructions regarding the aforementioned matters.";

            run363.Append(runProperties355);
            run363.Append(text214);

            paragraph92.Append(paragraphProperties88);
            paragraph92.Append(run349);
            paragraph92.Append(run350);
            paragraph92.Append(run351);
            paragraph92.Append(run352);
            paragraph92.Append(run353);
            paragraph92.Append(run354);
            paragraph92.Append(proofError5);
            paragraph92.Append(run355);
            paragraph92.Append(proofError6);
            paragraph92.Append(run356);
            paragraph92.Append(run357);
            paragraph92.Append(run358);
            paragraph92.Append(run359);
            paragraph92.Append(run360);
            paragraph92.Append(run361);
            paragraph92.Append(run362);
            paragraph92.Append(run363);

            Paragraph paragraph93 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F03C81", RsidParagraphProperties = "00F03C81", RsidRunAdditionDefault = "00F03C81", ParagraphId = "7EF5D7C1", TextId = "498C5C36" };

            ParagraphProperties paragraphProperties89 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation198 = new Indentation() { Start = "195", StartCharacters = 29, Hanging = "125", HangingChars = 78 };

            ParagraphMarkRunProperties paragraphMarkRunProperties85 = new ParagraphMarkRunProperties();
            FontSize fontSize384 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript120 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties85.Append(fontSize384);
            paragraphMarkRunProperties85.Append(fontSizeComplexScript120);

            paragraphProperties89.Append(spacingBetweenLines25);
            paragraphProperties89.Append(indentation198);
            paragraphProperties89.Append(paragraphMarkRunProperties85);

            Run run364 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties356 = new RunProperties();
            FontSize fontSize385 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript121 = new FontSizeComplexScript() { Val = "16" };

            runProperties356.Append(fontSize385);
            runProperties356.Append(fontSizeComplexScript121);
            Text text215 = new Text();
            text215.Text = "b)";

            run364.Append(runProperties356);
            run364.Append(text215);

            Run run365 = new Run() { RsidRunProperties = "000B6332", RsidRunAddition = "000B6332" };

            RunProperties runProperties357 = new RunProperties();
            FontSize fontSize386 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript122 = new FontSizeComplexScript() { Val = "16" };

            runProperties357.Append(fontSize386);
            runProperties357.Append(fontSizeComplexScript122);
            Text text216 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text216.Text = "Abnormal customer ";

            run365.Append(runProperties357);
            run365.Append(text216);

            Run run366 = new Run() { RsidRunProperties = "000B6332", RsidRunAddition = "00153D64" };

            RunProperties runProperties358 = new RunProperties();
            FontSize fontSize387 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript123 = new FontSizeComplexScript() { Val = "16" };

            runProperties358.Append(fontSize387);
            runProperties358.Append(fontSizeComplexScript123);
            Text text217 = new Text();
            text217.Text = "complaints";

            run366.Append(runProperties358);
            run366.Append(text217);

            Run run367 = new Run() { RsidRunProperties = "000B6332", RsidRunAddition = "000B6332" };

            RunProperties runProperties359 = new RunProperties();
            FontSize fontSize388 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript124 = new FontSizeComplexScript() { Val = "16" };

            runProperties359.Append(fontSize388);
            runProperties359.Append(fontSizeComplexScript124);
            Text text218 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text218.Text = " regarding structured products";

            run367.Append(runProperties359);
            run367.Append(text218);

            Run run368 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties360 = new RunProperties();
            FontSize fontSize389 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript125 = new FontSizeComplexScript() { Val = "16" };

            runProperties360.Append(fontSize389);
            runProperties360.Append(fontSizeComplexScript125);
            Text text219 = new Text();
            text219.Text = "：";

            run368.Append(runProperties360);
            run368.Append(text219);

            Run run369 = new Run() { RsidRunAddition = "000B6332" };

            RunProperties runProperties361 = new RunProperties();
            RunFonts runFonts450 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize390 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript126 = new FontSizeComplexScript() { Val = "16" };

            runProperties361.Append(runFonts450);
            runProperties361.Append(fontSize390);
            runProperties361.Append(fontSizeComplexScript126);
            Text text220 = new Text();
            text220.Text = "th";

            run369.Append(runProperties361);
            run369.Append(text220);

            Run run370 = new Run() { RsidRunAddition = "000B6332" };

            RunProperties runProperties362 = new RunProperties();
            FontSize fontSize391 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript127 = new FontSizeComplexScript() { Val = "16" };

            runProperties362.Append(fontSize391);
            runProperties362.Append(fontSizeComplexScript127);
            Text text221 = new Text();
            text221.Text = "is refers to FSC’s written instruction. (2017/6/5 No.";

            run370.Append(runProperties362);
            run370.Append(text221);

            Run run371 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties363 = new RunProperties();
            FontSize fontSize392 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript128 = new FontSizeComplexScript() { Val = "16" };

            runProperties363.Append(fontSize392);
            runProperties363.Append(fontSizeComplexScript128);
            Text text222 = new Text();
            text222.Text = "1060152199";

            run371.Append(runProperties363);
            run371.Append(text222);

            Run run372 = new Run() { RsidRunAddition = "000B6332" };

            RunProperties runProperties364 = new RunProperties();
            FontSize fontSize393 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript129 = new FontSizeComplexScript() { Val = "16" };

            runProperties364.Append(fontSize393);
            runProperties364.Append(fontSizeComplexScript129);
            Text text223 = new Text();
            text223.Text = ")";

            run372.Append(runProperties364);
            run372.Append(text223);

            paragraph93.Append(paragraphProperties89);
            paragraph93.Append(run364);
            paragraph93.Append(run365);
            paragraph93.Append(run366);
            paragraph93.Append(run367);
            paragraph93.Append(run368);
            paragraph93.Append(run369);
            paragraph93.Append(run370);
            paragraph93.Append(run371);
            paragraph93.Append(run372);

            Paragraph paragraph94 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F03C81", RsidParagraphProperties = "008410C6", RsidRunAdditionDefault = "00F03C81", ParagraphId = "38630DD6", TextId = "7474229B" };

            ParagraphProperties paragraphProperties90 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation199 = new Indentation() { Start = "195", StartCharacters = 29, Hanging = "125", HangingChars = 78 };

            paragraphProperties90.Append(spacingBetweenLines26);
            paragraphProperties90.Append(indentation199);

            Run run373 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties365 = new RunProperties();
            RunFonts runFonts451 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize394 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript130 = new FontSizeComplexScript() { Val = "16" };

            runProperties365.Append(runFonts451);
            runProperties365.Append(fontSize394);
            runProperties365.Append(fontSizeComplexScript130);
            Text text224 = new Text();
            text224.Text = "c)";

            run373.Append(runProperties365);
            run373.Append(text224);

            Run run374 = new Run() { RsidRunAddition = "004C41E7" };

            RunProperties runProperties366 = new RunProperties();
            FontSize fontSize395 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript131 = new FontSizeComplexScript() { Val = "16" };

            runProperties366.Append(fontSize395);
            runProperties366.Append(fontSizeComplexScript131);
            Text text225 = new Text();
            text225.Text = "I";

            run374.Append(runProperties366);
            run374.Append(text225);

            Run run375 = new Run() { RsidRunAddition = "00932310" };

            RunProperties runProperties367 = new RunProperties();
            FontSize fontSize396 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript132 = new FontSizeComplexScript() { Val = "16" };

            runProperties367.Append(fontSize396);
            runProperties367.Append(fontSizeComplexScript132);
            Text text226 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text226.Text = "t is advisable ";

            run375.Append(runProperties367);
            run375.Append(text226);

            Run run376 = new Run() { RsidRunAddition = "00C56711" };

            RunProperties runProperties368 = new RunProperties();
            FontSize fontSize397 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript133 = new FontSizeComplexScript() { Val = "16" };

            runProperties368.Append(fontSize397);
            runProperties368.Append(fontSizeComplexScript133);
            Text text227 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text227.Text = "to ";

            run376.Append(runProperties368);
            run376.Append(text227);

            Run run377 = new Run() { RsidRunAddition = "00932310" };

            RunProperties runProperties369 = new RunProperties();
            FontSize fontSize398 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript134 = new FontSizeComplexScript() { Val = "16" };

            runProperties369.Append(fontSize398);
            runProperties369.Append(fontSizeComplexScript134);
            Text text228 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text228.Text = "incorporate ";

            run377.Append(runProperties369);
            run377.Append(text228);

            Run run378 = new Run() { RsidRunAddition = "00C56711" };

            RunProperties runProperties370 = new RunProperties();
            FontSize fontSize399 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript135 = new FontSizeComplexScript() { Val = "16" };

            runProperties370.Append(fontSize399);
            runProperties370.Append(fontSizeComplexScript135);
            Text text229 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text229.Text = "any ";

            run378.Append(runProperties370);
            run378.Append(text229);

            Run run379 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties371 = new RunProperties();
            FontSize fontSize400 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript136 = new FontSizeComplexScript() { Val = "16" };

            runProperties371.Append(fontSize400);
            runProperties371.Append(fontSizeComplexScript136);
            Text text230 = new Text();
            text230.Text = "significant";

            run379.Append(runProperties371);
            run379.Append(text230);

            Run run380 = new Run() { RsidRunAddition = "00C56711" };

            RunProperties runProperties372 = new RunProperties();
            FontSize fontSize401 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript137 = new FontSizeComplexScript() { Val = "16" };

            runProperties372.Append(fontSize401);
            runProperties372.Append(fontSizeComplexScript137);
            Text text231 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text231.Text = " control deficiency identified ";

            run380.Append(runProperties372);
            run380.Append(text231);

            Run run381 = new Run() { RsidRunAddition = "00932310" };

            RunProperties runProperties373 = new RunProperties();
            FontSize fontSize402 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript138 = new FontSizeComplexScript() { Val = "16" };

            runProperties373.Append(fontSize402);
            runProperties373.Append(fontSizeComplexScript138);
            Text text232 = new Text();
            text232.Text = "in";

            run381.Append(runProperties373);
            run381.Append(text232);

            Run run382 = new Run() { RsidRunAddition = "00C56711" };

            RunProperties runProperties374 = new RunProperties();
            FontSize fontSize403 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript139 = new FontSizeComplexScript() { Val = "16" };

            runProperties374.Append(fontSize403);
            runProperties374.Append(fontSizeComplexScript139);
            Text text233 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text233.Text = " customer ";

            run382.Append(runProperties374);
            run382.Append(text233);

            Run run383 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties375 = new RunProperties();
            FontSize fontSize404 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript140 = new FontSizeComplexScript() { Val = "16" };

            runProperties375.Append(fontSize404);
            runProperties375.Append(fontSizeComplexScript140);
            Text text234 = new Text();
            text234.Text = "complaints";

            run383.Append(runProperties375);
            run383.Append(text234);

            Run run384 = new Run() { RsidRunAddition = "00C56711" };

            RunProperties runProperties376 = new RunProperties();
            FontSize fontSize405 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript141 = new FontSizeComplexScript() { Val = "16" };

            runProperties376.Append(fontSize405);
            runProperties376.Append(fontSizeComplexScript141);
            Text text235 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text235.Text = " or ";

            run384.Append(runProperties376);
            run384.Append(text235);

            Run run385 = new Run() { RsidRunAddition = "00932310" };

            RunProperties runProperties377 = new RunProperties();
            FontSize fontSize406 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript142 = new FontSizeComplexScript() { Val = "16" };

            runProperties377.Append(fontSize406);
            runProperties377.Append(fontSizeComplexScript142);
            Text text236 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text236.Text = "cases reported by ";

            run385.Append(runProperties377);
            run385.Append(text236);

            Run run386 = new Run() { RsidRunAddition = "00153D64" };

            RunProperties runProperties378 = new RunProperties();
            FontSize fontSize407 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript143 = new FontSizeComplexScript() { Val = "16" };

            runProperties378.Append(fontSize407);
            runProperties378.Append(fontSizeComplexScript143);
            Text text237 = new Text();
            text237.Text = "whistleblowers";

            run386.Append(runProperties378);
            run386.Append(text237);

            Run run387 = new Run() { RsidRunAddition = "004C41E7" };

            RunProperties runProperties379 = new RunProperties();
            FontSize fontSize408 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript144 = new FontSizeComplexScript() { Val = "16" };

            runProperties379.Append(fontSize408);
            runProperties379.Append(fontSizeComplexScript144);
            Text text238 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text238.Text = " into audit focus";

            run387.Append(runProperties379);
            run387.Append(text238);

            Run run388 = new Run() { RsidRunAddition = "00C56711" };

            RunProperties runProperties380 = new RunProperties();
            FontSize fontSize409 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript145 = new FontSizeComplexScript() { Val = "16" };

            runProperties380.Append(fontSize409);
            runProperties380.Append(fontSizeComplexScript145);
            Text text239 = new Text();
            text239.Text = ".";

            run388.Append(runProperties380);
            run388.Append(text239);

            paragraph94.Append(paragraphProperties90);
            paragraph94.Append(run373);
            paragraph94.Append(run374);
            paragraph94.Append(run375);
            paragraph94.Append(run376);
            paragraph94.Append(run377);
            paragraph94.Append(run378);
            paragraph94.Append(run379);
            paragraph94.Append(run380);
            paragraph94.Append(run381);
            paragraph94.Append(run382);
            paragraph94.Append(run383);
            paragraph94.Append(run384);
            paragraph94.Append(run385);
            paragraph94.Append(run386);
            paragraph94.Append(run387);
            paragraph94.Append(run388);

            footnote5.Append(paragraph92);
            footnote5.Append(paragraph93);
            footnote5.Append(paragraph94);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);
            footnotes1.Append(footnote3);
            footnotes1.Append(footnote4);
            footnotes1.Append(footnote5);

            footnotesPart1.Footnotes = footnotes1;
        }

        // Generates content of wordprocessingPeoplePart1.
        private void GenerateWordprocessingPeoplePart1Content(WordprocessingPeoplePart wordprocessingPeoplePart1)
        {
            W15.People people1 = new W15.People() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14" } };
            people1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            people1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            people1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            people1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            people1.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            people1.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            people1.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            people1.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            people1.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            people1.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            people1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            people1.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            people1.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            people1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            people1.AddNamespaceDeclaration("oel", "http://schemas.microsoft.com/office/2019/extlst");
            people1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            people1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            people1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            people1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            people1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            people1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            people1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            people1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            people1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            people1.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            people1.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            people1.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            people1.AddNamespaceDeclaration("w16sdtdh", "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash");
            people1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            people1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            people1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            people1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            people1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            W15.Person person1 = new W15.Person() { Author = "余亭妍" };
            W15.PresenceInfo presenceInfo1 = new W15.PresenceInfo() { ProviderId = "AD", UserId = "S::natasha.yu@newtype.com.tw::380cf427-be10-433a-92ea-dcab8dd7d9cd" };

            person1.Append(presenceInfo1);

            people1.Append(person1);

            wordprocessingPeoplePart1.People = people1;
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
            document.PackageProperties.Creator = "Z00002110";
            document.PackageProperties.Title = "內部稽核查核計劃(範本)";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Revision = "4";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2023-08-11T06:06:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2023-08-11T08:10:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "余亭妍";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2023-03-16T09:17:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }


    }
}
