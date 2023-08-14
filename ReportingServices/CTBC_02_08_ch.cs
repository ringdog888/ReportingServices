using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Ds = DocumentFormat.OpenXml.CustomXmlDataProperties;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using System;
using System.Data;


namespace CTBC_02_08_OPENXML
{
    public class GeneratedClass
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

            EndnotesPart endnotesPart1 = mainDocumentPart1.AddNewPart<EndnotesPart>("rId10");
            GenerateEndnotesPart1Content(endnotesPart1);

            CustomXmlPart customXmlPart4 = mainDocumentPart1.AddNewPart<CustomXmlPart>("application/xml", "rId4");
            GenerateCustomXmlPart4Content(customXmlPart4);

            CustomXmlPropertiesPart customXmlPropertiesPart4 = customXmlPart4.AddNewPart<CustomXmlPropertiesPart>("rId1");
            GenerateCustomXmlPropertiesPart4Content(customXmlPropertiesPart4);

            FootnotesPart footnotesPart1 = mainDocumentPart1.AddNewPart<FootnotesPart>("rId9");
            GenerateFootnotesPart1Content(footnotesPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId14");
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
            totalTime1.Text = "15";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "3";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "94";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "541";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "4";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "634";
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

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "7084CA69", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a3" };
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };
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

            Run run1 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
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
            text1.Text = "內部稽核查核計";

            run1.Append(runProperties1);
            run1.Append(text1);

            Run run2 = new Run() { RsidRunAddition = "00F64F71" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            Bold bold3 = new Bold();
            BoldComplexScript boldComplexScript3 = new BoldComplexScript();
            FontSize fontSize3 = new FontSize() { Val = "40" };
            Underline underline3 = new Underline() { Val = UnderlineValues.Single };

            runProperties2.Append(runFonts3);
            runProperties2.Append(bold3);
            runProperties2.Append(boldComplexScript3);
            runProperties2.Append(fontSize3);
            runProperties2.Append(underline3);
            Text text2 = new Text();
            text2.Text = "畫";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(run2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00C90DDE", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "315613B5", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "a3" };
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };
            Bold bold4 = new Bold();
            BoldComplexScript boldComplexScript4 = new BoldComplexScript();
            FontSize fontSize4 = new FontSize() { Val = "40" };
            Underline underline4 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties2.Append(runFonts4);
            paragraphMarkRunProperties2.Append(bold4);
            paragraphMarkRunProperties2.Append(boldComplexScript4);
            paragraphMarkRunProperties2.Append(fontSize4);
            paragraphMarkRunProperties2.Append(underline4);

            paragraphProperties2.Append(paragraphStyleId2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            paragraph2.Append(paragraphProperties2);

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "9781", Type = TableWidthUnitValues.Dxa };
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
            GridColumn gridColumn1 = new GridColumn() { Width = "1791" };
            GridColumn gridColumn2 = new GridColumn() { Width = "7990" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00C90DDE", RsidTableRowProperties = "003627CF", ParagraphId = "5DDE486F", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00C90DDE", RsidParagraphAddition = "00C90DDE", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "4E3EE071", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize5 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(fontSize5);

            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00C90DDE" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize6 = new FontSize() { Val = "28" };

            runProperties3.Append(runFonts6);
            runProperties3.Append(fontSize6);
            Text text3 = new Text();
            text3.Text = "查程名稱：";

            run3.Append(runProperties3);
            run3.Append(text3);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph3);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00C90DDE", RsidParagraphProperties = "00C90DDE", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "5D353926", TextId = "2CB04151" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };
            Indentation indentation1 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification3 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold5 = new Bold();
            Color color1 = new Color() { Val = "0000FF" };
            FontSize fontSize7 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties4.Append(runFonts7);
            paragraphMarkRunProperties4.Append(bold5);
            paragraphMarkRunProperties4.Append(color1);
            paragraphMarkRunProperties4.Append(fontSize7);

            paragraphProperties4.Append(snapToGrid1);
            paragraphProperties4.Append(indentation1);
            paragraphProperties4.Append(justification3);
            paragraphProperties4.Append(paragraphMarkRunProperties4);

            Run run4 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts8 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold6 = new Bold();
            Color color2 = new Color() { Val = "0000FF" };
            FontSize fontSize8 = new FontSize() { Val = "28" };

            runProperties4.Append(runFonts8);
            runProperties4.Append(bold6);
            runProperties4.Append(color2);
            runProperties4.Append(fontSize8);
            Text text4 = new Text();
            text4.Text = dt.Rows[0]["planname"].ToString();

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph4.Append(paragraphProperties4);
            paragraph4.Append(run4);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph4);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "003627CF", ParagraphId = "2904FBEA", TextId = "77777777" };

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "0014137D", ParagraphId = "338DBE2A", TextId = "77777777" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId4 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties5.Append(paragraphStyleId4);

            Run run5 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize9 = new FontSize() { Val = "28" };

            runProperties5.Append(runFonts9);
            runProperties5.Append(fontSize9);
            Text text5 = new Text();
            text5.Text = "受檢單位：";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph5.Append(paragraphProperties5);
            paragraph5.Append(run5);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph5);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties4.Append(tableCellWidth4);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "5FC5240D", TextId = "202E62A3" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };
            Indentation indentation2 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification4 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts10 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold7 = new Bold();
            Color color3 = new Color() { Val = "0000FF" };
            FontSize fontSize10 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties5.Append(runFonts10);
            paragraphMarkRunProperties5.Append(bold7);
            paragraphMarkRunProperties5.Append(color3);
            paragraphMarkRunProperties5.Append(fontSize10);

            paragraphProperties6.Append(snapToGrid2);
            paragraphProperties6.Append(indentation2);
            paragraphProperties6.Append(justification4);
            paragraphProperties6.Append(paragraphMarkRunProperties5);

            Run run6 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold8 = new Bold();
            Color color4 = new Color() { Val = "0000FF" };
            FontSize fontSize11 = new FontSize() { Val = "28" };

            runProperties6.Append(runFonts11);
            runProperties6.Append(bold8);
            runProperties6.Append(color4);
            runProperties6.Append(fontSize11);
            Text text6 = new Text();
            text6.Text = dt.Rows[0]["auditplandept"].ToString();

            run6.Append(runProperties6);
            run6.Append(text6);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run6);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph6);

            tableRow2.Append(tableCell3);
            tableRow2.Append(tableCell4);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "003627CF", ParagraphId = "7D34786B", TextId = "77777777" };

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties5.Append(tableCellWidth5);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "0014137D", ParagraphId = "3F40F6C8", TextId = "77777777" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId5 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            FontSize fontSize12 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties6.Append(fontSize12);

            paragraphProperties7.Append(paragraphStyleId5);
            paragraphProperties7.Append(paragraphMarkRunProperties6);

            Run run7 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize13 = new FontSize() { Val = "28" };

            runProperties7.Append(runFonts12);
            runProperties7.Append(fontSize13);
            Text text7 = new Text();
            text7.Text = "查核方式：";

            run7.Append(runProperties7);
            run7.Append(text7);

            paragraph7.Append(paragraphProperties7);
            paragraph7.Append(run7);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph7);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties6.Append(tableCellWidth6);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C90DDE", ParagraphId = "3A482807", TextId = "56518392" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            SnapToGrid snapToGrid3 = new SnapToGrid() { Val = false };
            Indentation indentation3 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification5 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold9 = new Bold();
            Color color5 = new Color() { Val = "0000FF" };
            FontSize fontSize14 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties7.Append(runFonts13);
            paragraphMarkRunProperties7.Append(bold9);
            paragraphMarkRunProperties7.Append(color5);
            paragraphMarkRunProperties7.Append(fontSize14);

            paragraphProperties8.Append(snapToGrid3);
            paragraphProperties8.Append(indentation3);
            paragraphProperties8.Append(justification5);
            paragraphProperties8.Append(paragraphMarkRunProperties7);

            Run run8 = new Run() { RsidRunProperties = "00C90DDE" };

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold10 = new Bold();
            Color color6 = new Color() { Val = "0000FF" };
            FontSize fontSize15 = new FontSize() { Val = "28" };

            runProperties8.Append(runFonts14);
            runProperties8.Append(bold10);
            runProperties8.Append(color6);
            runProperties8.Append(fontSize15);
            Text text8 = new Text();
            text8.Text = dt.Rows[0]["plantype"].ToString();

            run8.Append(runProperties8);
            run8.Append(text8);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run8);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph8);

            tableRow3.Append(tableCell5);
            tableRow3.Append(tableCell6);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "003627CF", ParagraphId = "0B27EEB0", TextId = "77777777" };

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties7.Append(tableCellWidth7);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "0014137D", ParagraphId = "60FCA5F1", TextId = "77777777" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId6 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            FontSize fontSize16 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties8.Append(fontSize16);

            paragraphProperties9.Append(paragraphStyleId6);
            paragraphProperties9.Append(paragraphMarkRunProperties8);

            Run run9 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize17 = new FontSize() { Val = "28" };

            runProperties9.Append(runFonts15);
            runProperties9.Append(fontSize17);
            Text text9 = new Text();
            text9.Text = "查核期間：";

            run9.Append(runProperties9);
            run9.Append(text9);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run9);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph9);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties8.Append(tableCellWidth8);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C7460E", ParagraphId = "5776FF2B", TextId = "232E6450" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();
            SnapToGrid snapToGrid4 = new SnapToGrid() { Val = false };
            Indentation indentation4 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification6 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts16 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold11 = new Bold();
            Color color7 = new Color() { Val = "0000FF" };
            FontSize fontSize18 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties9.Append(runFonts16);
            paragraphMarkRunProperties9.Append(bold11);
            paragraphMarkRunProperties9.Append(color7);
            paragraphMarkRunProperties9.Append(fontSize18);

            paragraphProperties10.Append(snapToGrid4);
            paragraphProperties10.Append(indentation4);
            paragraphProperties10.Append(justification6);
            paragraphProperties10.Append(paragraphMarkRunProperties9);

            Run run10 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold12 = new Bold();
            Color color8 = new Color() { Val = "0000FF" };
            FontSize fontSize19 = new FontSize() { Val = "28" };

            runProperties10.Append(runFonts17);
            runProperties10.Append(bold12);
            runProperties10.Append(color8);
            runProperties10.Append(fontSize19);
            Text text10 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["startdate"].ToString(), out DateTime date1))
            {
                text10.Text=date1.ToString("yyyy-mm-dd");
            }
            else
            {
                text10.Text = "";
            }


            run10.Append(runProperties10);
            run10.Append(text10);

            Run run11 = new Run() { RsidRunAddition = "008A0512" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold13 = new Bold();
            Color color9 = new Color() { Val = "0000FF" };
            FontSize fontSize20 = new FontSize() { Val = "28" };

            runProperties11.Append(runFonts18);
            runProperties11.Append(bold13);
            runProperties11.Append(color9);
            runProperties11.Append(fontSize20);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = " ";

            run11.Append(runProperties11);
            run11.Append(text11);

            Run run12 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold14 = new Bold();
            Color color10 = new Color() { Val = "0000FF" };
            FontSize fontSize21 = new FontSize() { Val = "28" };

            runProperties12.Append(runFonts19);
            runProperties12.Append(bold14);
            runProperties12.Append(color10);
            runProperties12.Append(fontSize21);
            Text text12 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text12.Text = "~ ";

            run12.Append(runProperties12);
            run12.Append(text12);

            Run run13 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold15 = new Bold();
            Color color11 = new Color() { Val = "0000FF" };
            FontSize fontSize22 = new FontSize() { Val = "28" };

            runProperties13.Append(runFonts20);
            runProperties13.Append(bold15);
            runProperties13.Append(color11);
            runProperties13.Append(fontSize22);
            Text text13 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["enddate"].ToString(), out DateTime date6))
            {
                text13.Text = date6.ToString("yyyy-mm-dd");
            }
            else
            {
                text13.Text = "";
            }

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph10.Append(paragraphProperties10);
            paragraph10.Append(run10);
            paragraph10.Append(run11);
            paragraph10.Append(run12);
            paragraph10.Append(run13);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph10);

            tableRow4.Append(tableCell7);
            tableRow4.Append(tableCell8);

            TableRow tableRow5 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "003627CF", ParagraphId = "5D680483", TextId = "77777777" };

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties9.Append(tableCellWidth9);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "007129F1", ParagraphId = "15FC9627", TextId = "77777777" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId7 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            FontSize fontSize23 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties10.Append(fontSize23);

            paragraphProperties11.Append(paragraphStyleId7);
            paragraphProperties11.Append(paragraphMarkRunProperties10);

            Run run14 = new Run() { RsidRunProperties = "00CD2680" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize24 = new FontSize() { Val = "28" };

            runProperties14.Append(runFonts21);
            runProperties14.Append(fontSize24);
            Text text14 = new Text();
            text14.Text = "查核範圍：";

            run14.Append(runProperties14);
            run14.Append(text14);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run14);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph11);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties10.Append(tableCellWidth10);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C7460E", ParagraphId = "16192332", TextId = "44AC9D93" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            SnapToGrid snapToGrid5 = new SnapToGrid() { Val = false };
            Indentation indentation5 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification7 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts22 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold16 = new Bold();
            Color color12 = new Color() { Val = "0000FF" };
            FontSize fontSize25 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties11.Append(runFonts22);
            paragraphMarkRunProperties11.Append(bold16);
            paragraphMarkRunProperties11.Append(color12);
            paragraphMarkRunProperties11.Append(fontSize25);

            paragraphProperties12.Append(snapToGrid5);
            paragraphProperties12.Append(indentation5);
            paragraphProperties12.Append(justification7);
            paragraphProperties12.Append(paragraphMarkRunProperties11);

            Run run15 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts23 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold17 = new Bold();
            Color color13 = new Color() { Val = "0000FF" };
            FontSize fontSize26 = new FontSize() { Val = "28" };

            runProperties15.Append(runFonts23);
            runProperties15.Append(bold17);
            runProperties15.Append(color13);
            runProperties15.Append(fontSize26);
            Text text15 = new Text();
            //text15.Text = "查核範圍起日";
            if (DateTime.TryParse(dt.Rows[0]["ar_startdate"].ToString(), out DateTime date7))
            {
                text15.Text = date7.ToString("yyyy-mm-dd");
            }
            else
            {
                text15.Text = "";
            }

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts24 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold18 = new Bold();
            Color color14 = new Color() { Val = "0000FF" };
            FontSize fontSize27 = new FontSize() { Val = "28" };

            runProperties16.Append(runFonts24);
            runProperties16.Append(bold18);
            runProperties16.Append(color14);
            runProperties16.Append(fontSize27);
            Text text16 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text16.Text = " ~ ";

            run16.Append(runProperties16);
            run16.Append(text16);

            Run run17 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold19 = new Bold();
            Color color15 = new Color() { Val = "0000FF" };
            FontSize fontSize28 = new FontSize() { Val = "28" };

            runProperties17.Append(runFonts25);
            runProperties17.Append(bold19);
            runProperties17.Append(color15);
            runProperties17.Append(fontSize28);
            Text text17 = new Text();
            //text17.Text = "查核範圍迄日";
            if (DateTime.TryParse(dt.Rows[0]["ar_enddate"].ToString(), out DateTime date8))
            {
                text17.Text = date7.ToString("yyyy-mm-dd");
            }
            else
            {
                text17.Text = "";
            }

            run17.Append(runProperties17);
            run17.Append(text17);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run15);
            paragraph12.Append(run16);
            paragraph12.Append(run17);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph12);

            tableRow5.Append(tableCell9);
            tableRow5.Append(tableCell10);

            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00C7460E", RsidTableRowProperties = "003627CF", ParagraphId = "1105724F", TextId = "77777777" };

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties11.Append(tableCellWidth11);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00C7460E", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00C7460E", ParagraphId = "7A81FD23", TextId = "77777777" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId8 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            FontSize fontSize29 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties12.Append(fontSize29);

            paragraphProperties13.Append(paragraphStyleId8);
            paragraphProperties13.Append(paragraphMarkRunProperties12);

            Run run18 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize30 = new FontSize() { Val = "28" };

            runProperties18.Append(runFonts26);
            runProperties18.Append(fontSize30);
            Text text18 = new Text();
            text18.Text = "領";

            run18.Append(runProperties18);
            run18.Append(text18);

            Run run19 = new Run() { RsidRunAddition = "000C4189" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize31 = new FontSize() { Val = "28" };

            runProperties19.Append(runFonts27);
            runProperties19.Append(fontSize31);
            Text text19 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text19.Text = "　　";

            run19.Append(runProperties19);
            run19.Append(text19);

            Run run20 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts28 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize32 = new FontSize() { Val = "28" };

            runProperties20.Append(runFonts28);
            runProperties20.Append(fontSize32);
            Text text20 = new Text();
            text20.Text = "隊：";

            run20.Append(runProperties20);
            run20.Append(text20);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run18);
            paragraph13.Append(run19);
            paragraph13.Append(run20);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph13);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties12.Append(tableCellWidth12);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "00C7460E", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "00C7460E", ParagraphId = "67D928F7", TextId = "14AFE363" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            SnapToGrid snapToGrid6 = new SnapToGrid() { Val = false };
            Indentation indentation6 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification8 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts29 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold20 = new Bold();
            Color color16 = new Color() { Val = "0000FF" };
            FontSize fontSize33 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties13.Append(runFonts29);
            paragraphMarkRunProperties13.Append(bold20);
            paragraphMarkRunProperties13.Append(color16);
            paragraphMarkRunProperties13.Append(fontSize33);

            paragraphProperties14.Append(snapToGrid6);
            paragraphProperties14.Append(indentation6);
            paragraphProperties14.Append(justification8);
            paragraphProperties14.Append(paragraphMarkRunProperties13);

            Run run21 = new Run();

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold21 = new Bold();
            Color color17 = new Color() { Val = "0000FF" };
            FontSize fontSize34 = new FontSize() { Val = "28" };

            runProperties21.Append(runFonts30);
            runProperties21.Append(bold21);
            runProperties21.Append(color17);
            runProperties21.Append(fontSize34);
            Text text21 = new Text();
            text21.Text = dt.Rows[0]["leader"].ToString();

            run21.Append(runProperties21);
            run21.Append(text21);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run21);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph14);

            tableRow6.Append(tableCell11);
            tableRow6.Append(tableCell12);

            TableRow tableRow7 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "003627CF", ParagraphId = "22A589ED", TextId = "77777777" };

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties13.Append(tableCellWidth13);

            Paragraph paragraph15 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "0014137D", ParagraphId = "18FF0DBA", TextId = "77777777" };

            ParagraphProperties paragraphProperties15 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId9 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            FontSize fontSize35 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties14.Append(fontSize35);

            paragraphProperties15.Append(paragraphStyleId9);
            paragraphProperties15.Append(paragraphMarkRunProperties14);

            Run run22 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize36 = new FontSize() { Val = "28" };

            runProperties22.Append(runFonts31);
            runProperties22.Append(fontSize36);
            Text text22 = new Text();
            text22.Text = "查核人員：";

            run22.Append(runProperties22);
            run22.Append(text22);

            paragraph15.Append(paragraphProperties15);
            paragraph15.Append(run22);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph15);

            TableCell tableCell14 = new TableCell();

            TableCellProperties tableCellProperties14 = new TableCellProperties();
            TableCellWidth tableCellWidth14 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties14.Append(tableCellWidth14);

            Paragraph paragraph16 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "003627CF", RsidRunAdditionDefault = "008625F1", ParagraphId = "0BBE67EC", TextId = "485F81E4" };

            ParagraphProperties paragraphProperties16 = new ParagraphProperties();
            SnapToGrid snapToGrid7 = new SnapToGrid() { Val = false };
            Indentation indentation7 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification9 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties15 = new ParagraphMarkRunProperties();
            RunFonts runFonts32 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold22 = new Bold();
            Color color18 = new Color() { Val = "0000FF" };
            FontSize fontSize37 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties15.Append(runFonts32);
            paragraphMarkRunProperties15.Append(bold22);
            paragraphMarkRunProperties15.Append(color18);
            paragraphMarkRunProperties15.Append(fontSize37);

            paragraphProperties16.Append(snapToGrid7);
            paragraphProperties16.Append(indentation7);
            paragraphProperties16.Append(justification9);
            paragraphProperties16.Append(paragraphMarkRunProperties15);

            Run run23 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold23 = new Bold();
            Color color19 = new Color() { Val = "0000FF" };
            FontSize fontSize38 = new FontSize() { Val = "28" };

            runProperties23.Append(runFonts33);
            runProperties23.Append(bold23);
            runProperties23.Append(color19);
            runProperties23.Append(fontSize38);
            Text text23 = new Text();
            text23.Text = dt.Rows[0]["Member"].ToString();

            run23.Append(runProperties23);
            run23.Append(text23);

            paragraph16.Append(paragraphProperties16);
            paragraph16.Append(run23);

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
            Paragraph paragraph17 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "43A8FD9F", TextId = "77777777" };

            Paragraph paragraph18 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "4516A343", TextId = "77777777" };

            ParagraphProperties paragraphProperties17 = new ParagraphProperties();
            SnapToGrid snapToGrid8 = new SnapToGrid() { Val = false };

            ParagraphMarkRunProperties paragraphMarkRunProperties16 = new ParagraphMarkRunProperties();
            Bold bold24 = new Bold();
            BoldComplexScript boldComplexScript5 = new BoldComplexScript();
            FontSize fontSize39 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties16.Append(bold24);
            paragraphMarkRunProperties16.Append(boldComplexScript5);
            paragraphMarkRunProperties16.Append(fontSize39);

            paragraphProperties17.Append(snapToGrid8);
            paragraphProperties17.Append(paragraphMarkRunProperties16);

            Run run24 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold25 = new Bold();
            BoldComplexScript boldComplexScript6 = new BoldComplexScript();
            FontSize fontSize40 = new FontSize() { Val = "28" };

            runProperties24.Append(runFonts34);
            runProperties24.Append(bold25);
            runProperties24.Append(boldComplexScript6);
            runProperties24.Append(fontSize40);
            Text text24 = new Text();
            text24.Text = "一、工作分配：";

            run24.Append(runProperties24);
            run24.Append(text24);

            paragraph18.Append(paragraphProperties17);
            paragraph18.Append(run24);

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

            TableRow tableRow8 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00193FEF", ParagraphId = "0ECB65DD", TextId = "77777777" };

            TableCell tableCell15 = new TableCell();

            TableCellProperties tableCellProperties15 = new TableCellProperties();
            TableCellWidth tableCellWidth15 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };

            tableCellProperties15.Append(tableCellWidth15);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "0014137D", ParagraphId = "10822DC4", TextId = "77777777" };

            ParagraphProperties paragraphProperties18 = new ParagraphProperties();
            SnapToGrid snapToGrid9 = new SnapToGrid() { Val = false };
            Indentation indentation8 = new Indentation() { Start = "-24", StartCharacters = -11, Hanging = "2" };
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties17 = new ParagraphMarkRunProperties();
            FontSize fontSize41 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties17.Append(fontSize41);

            paragraphProperties18.Append(snapToGrid9);
            paragraphProperties18.Append(indentation8);
            paragraphProperties18.Append(justification10);
            paragraphProperties18.Append(paragraphMarkRunProperties17);

            Run run25 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize42 = new FontSize() { Val = "28" };

            runProperties25.Append(runFonts35);
            runProperties25.Append(fontSize42);
            Text text25 = new Text();
            text25.Text = "編號";

            run25.Append(runProperties25);
            run25.Append(text25);

            paragraph19.Append(paragraphProperties18);
            paragraph19.Append(run25);

            tableCell15.Append(tableCellProperties15);
            tableCell15.Append(paragraph19);

            TableCell tableCell16 = new TableCell();

            TableCellProperties tableCellProperties16 = new TableCellProperties();
            TableCellWidth tableCellWidth16 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };

            tableCellProperties16.Append(tableCellWidth16);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "0014137D", ParagraphId = "6D22654F", TextId = "77777777" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            SnapToGrid snapToGrid10 = new SnapToGrid() { Val = false };
            Indentation indentation9 = new Indentation() { FirstLine = "280", FirstLineChars = 100 };
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties18 = new ParagraphMarkRunProperties();
            FontSize fontSize43 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties18.Append(fontSize43);

            paragraphProperties19.Append(snapToGrid10);
            paragraphProperties19.Append(indentation9);
            paragraphProperties19.Append(justification11);
            paragraphProperties19.Append(paragraphMarkRunProperties18);

            Run run26 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts36 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize44 = new FontSize() { Val = "28" };

            runProperties26.Append(runFonts36);
            runProperties26.Append(fontSize44);
            Text text26 = new Text();
            text26.Text = "科";

            run26.Append(runProperties26);
            run26.Append(text26);

            Run run27 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize45 = new FontSize() { Val = "28" };

            runProperties27.Append(runFonts37);
            runProperties27.Append(fontSize45);
            Text text27 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text27.Text = "   ";

            run27.Append(runProperties27);
            run27.Append(text27);

            Run run28 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize46 = new FontSize() { Val = "28" };

            runProperties28.Append(runFonts38);
            runProperties28.Append(fontSize46);
            Text text28 = new Text();
            text28.Text = "目";

            run28.Append(runProperties28);
            run28.Append(text28);

            paragraph20.Append(paragraphProperties19);
            paragraph20.Append(run26);
            paragraph20.Append(run27);
            paragraph20.Append(run28);

            tableCell16.Append(tableCellProperties16);
            tableCell16.Append(paragraph20);

            TableCell tableCell17 = new TableCell();

            TableCellProperties tableCellProperties17 = new TableCellProperties();
            TableCellWidth tableCellWidth17 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };

            tableCellProperties17.Append(tableCellWidth17);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "0014137D", ParagraphId = "5B956EE0", TextId = "77777777" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();
            SnapToGrid snapToGrid11 = new SnapToGrid() { Val = false };
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            FontSize fontSize47 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties19.Append(fontSize47);

            paragraphProperties20.Append(snapToGrid11);
            paragraphProperties20.Append(justification12);
            paragraphProperties20.Append(paragraphMarkRunProperties19);

            Run run29 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize48 = new FontSize() { Val = "28" };

            runProperties29.Append(runFonts39);
            runProperties29.Append(fontSize48);
            Text text29 = new Text();
            text29.Text = "查核人員";

            run29.Append(runProperties29);
            run29.Append(text29);

            paragraph21.Append(paragraphProperties20);
            paragraph21.Append(run29);

            tableCell17.Append(tableCellProperties17);
            tableCell17.Append(paragraph21);

            tableRow8.Append(tableCell15);
            tableRow8.Append(tableCell16);
            tableRow8.Append(tableCell17);

            TableRow tableRow9 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "0014137D", RsidTableRowProperties = "00193FEF", ParagraphId = "0DCFF67B", TextId = "77777777" };

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(tableCellVerticalAlignment1);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00DB5532", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "2C0F84A3", TextId = "74051F72" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            SnapToGrid snapToGrid12 = new SnapToGrid() { Val = false };
            Justification justification13 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts40 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };

            paragraphMarkRunProperties20.Append(runFonts40);

            paragraphProperties21.Append(snapToGrid12);
            paragraphProperties21.Append(justification13);
            paragraphProperties21.Append(paragraphMarkRunProperties20);

            Run run30 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties30 = new RunProperties();
            RunStyle runStyle1 = new RunStyle() { Val = "normaltextrun" };
            Bold bold26 = new Bold();
            BoldComplexScript boldComplexScript7 = new BoldComplexScript();
            Color color20 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "28" };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages1 = new Languages() { EastAsia = "zh-HK" };

            runProperties30.Append(runStyle1);
            runProperties30.Append(bold26);
            runProperties30.Append(boldComplexScript7);
            runProperties30.Append(color20);
            runProperties30.Append(fontSizeComplexScript1);
            runProperties30.Append(shading1);
            runProperties30.Append(languages1);
            Text text30 = new Text();
            text30.Text = "科目編號";
            //text30.Text = dt.Rows[0]["subcode"].ToString();

            run30.Append(runProperties30);
            run30.Append(text30);

            Run run31 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties31 = new RunProperties();
            RunStyle runStyle2 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts41 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold27 = new Bold();
            BoldComplexScript boldComplexScript8 = new BoldComplexScript();
            Color color21 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "28" };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages2 = new Languages() { EastAsia = "zh-HK" };

            runProperties31.Append(runStyle2);
            runProperties31.Append(runFonts41);
            runProperties31.Append(bold27);
            runProperties31.Append(boldComplexScript8);
            runProperties31.Append(color21);
            runProperties31.Append(fontSizeComplexScript2);
            runProperties31.Append(shading2);
            runProperties31.Append(languages2);
            Text text31 = new Text();
            text31.Text = "";

            run31.Append(runProperties31);
            run31.Append(text31);

            paragraph22.Append(paragraphProperties21);
            paragraph22.Append(run30);
            paragraph22.Append(run31);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph22);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(tableCellVerticalAlignment2);

            Paragraph paragraph23 = new Paragraph() { RsidParagraphMarkRevision = "00DB5532", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "6A602147", TextId = "604CB3BA" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            SnapToGrid snapToGrid13 = new SnapToGrid() { Val = false };
            Indentation indentation10 = new Indentation() { FirstLine = "240", FirstLineChars = 100 };
            Justification justification14 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts42 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };

            paragraphMarkRunProperties21.Append(runFonts42);

            paragraphProperties22.Append(snapToGrid13);
            paragraphProperties22.Append(indentation10);
            paragraphProperties22.Append(justification14);
            paragraphProperties22.Append(paragraphMarkRunProperties21);

            Run run32 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties32 = new RunProperties();
            RunStyle runStyle3 = new RunStyle() { Val = "normaltextrun" };
            Bold bold28 = new Bold();
            BoldComplexScript boldComplexScript9 = new BoldComplexScript();
            Color color22 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "28" };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages3 = new Languages() { EastAsia = "zh-HK" };

            runProperties32.Append(runStyle3);
            runProperties32.Append(bold28);
            runProperties32.Append(boldComplexScript9);
            runProperties32.Append(color22);
            runProperties32.Append(fontSizeComplexScript3);
            runProperties32.Append(shading3);
            runProperties32.Append(languages3);
            Text text32 = new Text();
            //text32.Text = "查核科目";
            text32.Text = dt.Rows[0]["L1Name"].ToString();

            run32.Append(runProperties32);
            run32.Append(text32);

            paragraph23.Append(paragraphProperties22);
            paragraph23.Append(run32);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph23);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment3 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(tableCellVerticalAlignment3);

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "00DB5532", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "04E98640", TextId = "321A846E" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            SnapToGrid snapToGrid14 = new SnapToGrid() { Val = false };
            Justification justification15 = new Justification() { Val = JustificationValues.Both };

            paragraphProperties23.Append(snapToGrid14);
            paragraphProperties23.Append(justification15);

            Run run33 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold29 = new Bold();
            Color color23 = new Color() { Val = "0000FF" };

            runProperties33.Append(runFonts43);
            runProperties33.Append(bold29);
            runProperties33.Append(color23);
            Text text33 = new Text();
            text33.Text = dt.Rows[0]["Member"].ToString();

            run33.Append(runProperties33);
            run33.Append(text33);

            paragraph24.Append(paragraphProperties23);
            paragraph24.Append(run33);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph24);

            tableRow9.Append(tableCell18);
            tableRow9.Append(tableCell19);
            tableRow9.Append(tableCell20);

            TableRow tableRow10 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00EB5D9F", RsidTableRowProperties = "00193FEF", ParagraphId = "319AD75A", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)80U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "1881", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment4 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(tableCellVerticalAlignment4);

            Paragraph paragraph25 = new Paragraph() { RsidParagraphMarkRevision = "00DB5532", RsidParagraphAddition = "00EB5D9F", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "7E2F8D1E", TextId = "03943BFF" };

            ParagraphProperties paragraphProperties24 = new ParagraphProperties();
            SnapToGrid snapToGrid15 = new SnapToGrid() { Val = false };
            Justification justification16 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };

            paragraphMarkRunProperties22.Append(runFonts44);

            paragraphProperties24.Append(snapToGrid15);
            paragraphProperties24.Append(justification16);
            paragraphProperties24.Append(paragraphMarkRunProperties22);

            Run run34 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties34 = new RunProperties();
            RunStyle runStyle4 = new RunStyle() { Val = "normaltextrun" };
            Bold bold30 = new Bold();
            BoldComplexScript boldComplexScript10 = new BoldComplexScript();
            Color color24 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "28" };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages4 = new Languages() { EastAsia = "zh-HK" };

            runProperties34.Append(runStyle4);
            runProperties34.Append(bold30);
            runProperties34.Append(boldComplexScript10);
            runProperties34.Append(color24);
            runProperties34.Append(fontSizeComplexScript4);
            runProperties34.Append(shading4);
            runProperties34.Append(languages4);
            Text text34 = new Text();
            //text34.Text = dt.Rows[0]["subcode"].ToString();

            run34.Append(runProperties34);
            run34.Append(text34);

            Run run35 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties35 = new RunProperties();
            RunStyle runStyle5 = new RunStyle() { Val = "normaltextrun" };
            RunFonts runFonts45 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold31 = new Bold();
            BoldComplexScript boldComplexScript11 = new BoldComplexScript();
            Color color25 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "28" };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages5 = new Languages() { EastAsia = "zh-HK" };

            runProperties35.Append(runStyle5);
            runProperties35.Append(runFonts45);
            runProperties35.Append(bold31);
            runProperties35.Append(boldComplexScript11);
            runProperties35.Append(color25);
            runProperties35.Append(fontSizeComplexScript5);
            runProperties35.Append(shading5);
            runProperties35.Append(languages5);
            Text text35 = new Text();
            text35.Text = "";

            run35.Append(runProperties35);
            run35.Append(text35);

            paragraph25.Append(paragraphProperties24);
            paragraph25.Append(run34);
            paragraph25.Append(run35);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph25);

            TableCell tableCell22 = new TableCell();

            TableCellProperties tableCellProperties22 = new TableCellProperties();
            TableCellWidth tableCellWidth22 = new TableCellWidth() { Width = "4425", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment5 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties22.Append(tableCellWidth22);
            tableCellProperties22.Append(tableCellVerticalAlignment5);

            Paragraph paragraph26 = new Paragraph() { RsidParagraphMarkRevision = "00DB5532", RsidParagraphAddition = "00EB5D9F", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "0145C032", TextId = "4FE7883D" };

            ParagraphProperties paragraphProperties25 = new ParagraphProperties();
            SnapToGrid snapToGrid16 = new SnapToGrid() { Val = false };
            Indentation indentation11 = new Indentation() { FirstLine = "240", FirstLineChars = 100 };
            Justification justification17 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            RunFonts runFonts46 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };

            paragraphMarkRunProperties23.Append(runFonts46);

            paragraphProperties25.Append(snapToGrid16);
            paragraphProperties25.Append(indentation11);
            paragraphProperties25.Append(justification17);
            paragraphProperties25.Append(paragraphMarkRunProperties23);

            Run run36 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties36 = new RunProperties();
            RunStyle runStyle6 = new RunStyle() { Val = "normaltextrun" };
            Bold bold32 = new Bold();
            BoldComplexScript boldComplexScript12 = new BoldComplexScript();
            Color color26 = new Color() { Val = "0000FF" };
            FontSizeComplexScript fontSizeComplexScript6 = new FontSizeComplexScript() { Val = "28" };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FFFFFF" };
            Languages languages6 = new Languages() { EastAsia = "zh-HK" };

            runProperties36.Append(runStyle6);
            runProperties36.Append(bold32);
            runProperties36.Append(boldComplexScript12);
            runProperties36.Append(color26);
            runProperties36.Append(fontSizeComplexScript6);
            runProperties36.Append(shading6);
            runProperties36.Append(languages6);
            Text text36 = new Text();
            //text36.Text = "查核科目";
            text36.Text = dt.Rows[0]["L1Name"].ToString();

            run36.Append(runProperties36);
            run36.Append(text36);

            paragraph26.Append(paragraphProperties25);
            paragraph26.Append(run36);

            tableCell22.Append(tableCellProperties22);
            tableCell22.Append(paragraph26);

            TableCell tableCell23 = new TableCell();

            TableCellProperties tableCellProperties23 = new TableCellProperties();
            TableCellWidth tableCellWidth23 = new TableCellWidth() { Width = "2908", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment6 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties23.Append(tableCellWidth23);
            tableCellProperties23.Append(tableCellVerticalAlignment6);

            Paragraph paragraph27 = new Paragraph() { RsidParagraphMarkRevision = "00DB5532", RsidParagraphAddition = "00EB5D9F", RsidParagraphProperties = "00193FEF", RsidRunAdditionDefault = "00EB5D9F", ParagraphId = "6BEB0139", TextId = "584E0BE0" };

            ParagraphProperties paragraphProperties26 = new ParagraphProperties();
            SnapToGrid snapToGrid17 = new SnapToGrid() { Val = false };
            Justification justification18 = new Justification() { Val = JustificationValues.Both };

            paragraphProperties26.Append(snapToGrid17);
            paragraphProperties26.Append(justification18);

            Run run37 = new Run() { RsidRunProperties = "00DB5532" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold33 = new Bold();
            Color color27 = new Color() { Val = "0000FF" };

            runProperties37.Append(runFonts47);
            runProperties37.Append(bold33);
            runProperties37.Append(color27);
            Text text37 = new Text();
            text37.Text = dt.Rows[0]["Member"].ToString();

            run37.Append(runProperties37);
            run37.Append(text37);

            paragraph27.Append(paragraphProperties26);
            paragraph27.Append(run37);

            tableCell23.Append(tableCellProperties23);
            tableCell23.Append(paragraph27);

            tableRow10.Append(tableRowProperties1);
            tableRow10.Append(tableCell21);
            tableRow10.Append(tableCell22);
            tableRow10.Append(tableCell23);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow8);
            table2.Append(tableRow9);
            table2.Append(tableRow10);
            Paragraph paragraph28 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00405678", RsidParagraphProperties = "000E3880", RsidRunAdditionDefault = "00405678", ParagraphId = "51EB9036", TextId = "77777777" };

            Paragraph paragraph29 = new Paragraph() { RsidParagraphMarkRevision = "00FE4552", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "001B4FF4", RsidRunAdditionDefault = "0014137D", ParagraphId = "55085B88", TextId = "77777777" };

            ParagraphProperties paragraphProperties27 = new ParagraphProperties();
            SnapToGrid snapToGrid18 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties24 = new ParagraphMarkRunProperties();
            Bold bold34 = new Bold();
            BoldComplexScript boldComplexScript13 = new BoldComplexScript();
            FontSize fontSize49 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties24.Append(bold34);
            paragraphMarkRunProperties24.Append(boldComplexScript13);
            paragraphMarkRunProperties24.Append(fontSize49);

            paragraphProperties27.Append(snapToGrid18);
            paragraphProperties27.Append(spacingBetweenLines1);
            paragraphProperties27.Append(paragraphMarkRunProperties24);

            Run run38 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties38 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold35 = new Bold();
            BoldComplexScript boldComplexScript14 = new BoldComplexScript();
            FontSize fontSize50 = new FontSize() { Val = "28" };

            runProperties38.Append(runFonts48);
            runProperties38.Append(bold35);
            runProperties38.Append(boldComplexScript14);
            runProperties38.Append(fontSize50);
            Text text38 = new Text();
            text38.Text = "二、該單位基本資料：";

            run38.Append(runProperties38);
            run38.Append(text38);

            Run run39 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties39 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold36 = new Bold();
            BoldComplexScript boldComplexScript15 = new BoldComplexScript();
            FontSize fontSize51 = new FontSize() { Val = "28" };

            runProperties39.Append(runFonts49);
            runProperties39.Append(bold36);
            runProperties39.Append(boldComplexScript15);
            runProperties39.Append(fontSize51);
            Text text39 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text39.Text = "  ";

            run39.Append(runProperties39);
            run39.Append(text39);

            paragraph29.Append(paragraphProperties27);
            paragraph29.Append(run38);
            paragraph29.Append(run39);

            Paragraph paragraph30 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00F55A18", RsidRunAdditionDefault = "0014137D", ParagraphId = "3CAC4A2C", TextId = "77777777" };

            ParagraphProperties paragraphProperties28 = new ParagraphProperties();
            SnapToGrid snapToGrid19 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties25 = new ParagraphMarkRunProperties();
            Bold bold37 = new Bold();
            BoldComplexScript boldComplexScript16 = new BoldComplexScript();
            FontSize fontSize52 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties25.Append(bold37);
            paragraphMarkRunProperties25.Append(boldComplexScript16);
            paragraphMarkRunProperties25.Append(fontSize52);

            paragraphProperties28.Append(snapToGrid19);
            paragraphProperties28.Append(spacingBetweenLines2);
            paragraphProperties28.Append(paragraphMarkRunProperties25);

            Run run40 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties40 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold38 = new Bold();
            BoldComplexScript boldComplexScript17 = new BoldComplexScript();
            FontSize fontSize53 = new FontSize() { Val = "28" };

            runProperties40.Append(runFonts50);
            runProperties40.Append(bold38);
            runProperties40.Append(boldComplexScript17);
            runProperties40.Append(fontSize53);
            Text text40 = new Text();
            text40.Text = "三、前次查核後組織、業務面或外部法規之變化：";

            run40.Append(runProperties40);
            run40.Append(text40);

            paragraph30.Append(paragraphProperties28);
            paragraph30.Append(run40);

            Paragraph paragraph31 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "38A644B0", TextId = "77777777" };

            ParagraphProperties paragraphProperties29 = new ParagraphProperties();

            NumberingProperties numberingProperties1 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference1 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId1 = new NumberingId() { Val = 1 };

            numberingProperties1.Append(numberingLevelReference1);
            numberingProperties1.Append(numberingId1);
            SnapToGrid snapToGrid20 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties26 = new ParagraphMarkRunProperties();
            FontSize fontSize54 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties26.Append(fontSize54);

            paragraphProperties29.Append(numberingProperties1);
            paragraphProperties29.Append(snapToGrid20);
            paragraphProperties29.Append(spacingBetweenLines3);
            paragraphProperties29.Append(paragraphMarkRunProperties26);

            Run run41 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties41 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize55 = new FontSize() { Val = "28" };

            runProperties41.Append(runFonts51);
            runProperties41.Append(fontSize55);
            Text text41 = new Text();
            text41.Text = "組織面：";

            run41.Append(runProperties41);
            run41.Append(text41);

            paragraph31.Append(paragraphProperties29);
            paragraph31.Append(run41);

            Paragraph paragraph32 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "5DA5F172", TextId = "77777777" };

            ParagraphProperties paragraphProperties30 = new ParagraphProperties();

            NumberingProperties numberingProperties2 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference2 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId2 = new NumberingId() { Val = 1 };

            numberingProperties2.Append(numberingLevelReference2);
            numberingProperties2.Append(numberingId2);
            SnapToGrid snapToGrid21 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines4 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties27 = new ParagraphMarkRunProperties();
            FontSize fontSize56 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties27.Append(fontSize56);

            paragraphProperties30.Append(numberingProperties2);
            paragraphProperties30.Append(snapToGrid21);
            paragraphProperties30.Append(spacingBetweenLines4);
            paragraphProperties30.Append(paragraphMarkRunProperties27);

            Run run42 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties42 = new RunProperties();
            RunFonts runFonts52 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize57 = new FontSize() { Val = "28" };

            runProperties42.Append(runFonts52);
            runProperties42.Append(fontSize57);
            Text text42 = new Text();
            text42.Text = "業務面：";

            run42.Append(runProperties42);
            run42.Append(text42);

            paragraph32.Append(paragraphProperties30);
            paragraph32.Append(run42);

            Paragraph paragraph33 = new Paragraph() { RsidParagraphMarkRevision = "00FE4552", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "0014137D", RsidRunAdditionDefault = "0014137D", ParagraphId = "422C96A1", TextId = "77777777" };

            ParagraphProperties paragraphProperties31 = new ParagraphProperties();

            NumberingProperties numberingProperties3 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference3 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId3 = new NumberingId() { Val = 1 };

            numberingProperties3.Append(numberingLevelReference3);
            numberingProperties3.Append(numberingId3);
            SnapToGrid snapToGrid22 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines5 = new SpacingBetweenLines() { Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties28 = new ParagraphMarkRunProperties();
            FontSize fontSize58 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties28.Append(fontSize58);

            paragraphProperties31.Append(numberingProperties3);
            paragraphProperties31.Append(snapToGrid22);
            paragraphProperties31.Append(spacingBetweenLines5);
            paragraphProperties31.Append(paragraphMarkRunProperties28);

            Run run43 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts53 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize59 = new FontSize() { Val = "28" };

            runProperties43.Append(runFonts53);
            runProperties43.Append(fontSize59);
            Text text43 = new Text();
            text43.Text = "外部法規";

            run43.Append(runProperties43);
            run43.Append(text43);

            Run run44 = new Run() { RsidRunAddition = "006C3779" };

            RunProperties runProperties44 = new RunProperties();
            RunStyle runStyle7 = new RunStyle() { Val = "a9" };
            FontSize fontSize60 = new FontSize() { Val = "28" };

            runProperties44.Append(runStyle7);
            runProperties44.Append(fontSize60);
            FootnoteReference footnoteReference1 = new FootnoteReference() { Id = 1 };

            run44.Append(runProperties44);
            run44.Append(footnoteReference1);

            Run run45 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties45 = new RunProperties();
            RunFonts runFonts54 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize61 = new FontSize() { Val = "28" };

            runProperties45.Append(runFonts54);
            runProperties45.Append(fontSize61);
            Text text44 = new Text();
            text44.Text = "：";

            run45.Append(runProperties45);
            run45.Append(text44);

            paragraph33.Append(paragraphProperties31);
            paragraph33.Append(run43);
            paragraph33.Append(run44);
            paragraph33.Append(run45);

            Paragraph paragraph34 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "0014137D", RsidParagraphProperties = "00F55A18", RsidRunAdditionDefault = "0014137D", ParagraphId = "439B6DC3", TextId = "77777777" };

            ParagraphProperties paragraphProperties32 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines6 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "0", LineRule = LineSpacingRuleValues.AtLeast };

            ParagraphMarkRunProperties paragraphMarkRunProperties29 = new ParagraphMarkRunProperties();
            RunFonts runFonts55 = new RunFonts() { Ascii = "標楷體" };
            Kern kern1 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize62 = new FontSize() { Val = "28" };
            FontSizeComplexScript fontSizeComplexScript7 = new FontSizeComplexScript() { Val = "23" };

            paragraphMarkRunProperties29.Append(runFonts55);
            paragraphMarkRunProperties29.Append(kern1);
            paragraphMarkRunProperties29.Append(fontSize62);
            paragraphMarkRunProperties29.Append(fontSizeComplexScript7);

            paragraphProperties32.Append(spacingBetweenLines6);
            paragraphProperties32.Append(paragraphMarkRunProperties29);

            Run run46 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts56 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold39 = new Bold();
            BoldComplexScript boldComplexScript18 = new BoldComplexScript();
            FontSize fontSize63 = new FontSize() { Val = "28" };

            runProperties46.Append(runFonts56);
            runProperties46.Append(bold39);
            runProperties46.Append(boldComplexScript18);
            runProperties46.Append(fontSize63);
            Text text45 = new Text();
            text45.Text = "四、前次查核重大缺失";

            run46.Append(runProperties46);
            run46.Append(text45);

            Run run47 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00882386" };

            RunProperties runProperties47 = new RunProperties();
            RunFonts runFonts57 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            Bold bold40 = new Bold();
            BoldComplexScript boldComplexScript19 = new BoldComplexScript();

            runProperties47.Append(runFonts57);
            runProperties47.Append(bold40);
            runProperties47.Append(boldComplexScript19);
            Text text46 = new Text();
            text46.Text = "（";

            run47.Append(runProperties47);
            run47.Append(text46);

            Run run48 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00882386" };

            RunProperties runProperties48 = new RunProperties();
            RunFonts runFonts58 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold41 = new Bold();
            BoldComplexScript boldComplexScript20 = new BoldComplexScript();

            runProperties48.Append(runFonts58);
            runProperties48.Append(bold41);
            runProperties48.Append(boldComplexScript20);
            Text text47 = new Text();
            text47.Text = "包含內、外部查核缺失";

            run48.Append(runProperties48);
            run48.Append(text47);

            Run run49 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00882386" };

            RunProperties runProperties49 = new RunProperties();
            RunFonts runFonts59 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            Bold bold42 = new Bold();
            BoldComplexScript boldComplexScript21 = new BoldComplexScript();

            runProperties49.Append(runFonts59);
            runProperties49.Append(bold42);
            runProperties49.Append(boldComplexScript21);
            Text text48 = new Text();
            text48.Text = "）";

            run49.Append(runProperties49);
            run49.Append(text48);

            Run run50 = new Run() { RsidRunProperties = "005F213D", RsidRunAddition = "00CF43E8" };

            RunProperties runProperties50 = new RunProperties();
            RunStyle runStyle8 = new RunStyle() { Val = "a9" };
            Bold bold43 = new Bold();
            BoldComplexScript boldComplexScript22 = new BoldComplexScript();
            Color color28 = new Color() { Val = "000000" };
            FontSize fontSize64 = new FontSize() { Val = "28" };

            runProperties50.Append(runStyle8);
            runProperties50.Append(bold43);
            runProperties50.Append(boldComplexScript22);
            runProperties50.Append(color28);
            runProperties50.Append(fontSize64);
            FootnoteReference footnoteReference2 = new FootnoteReference() { Id = 2 };

            run50.Append(runProperties50);
            run50.Append(footnoteReference2);

            Run run51 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties51 = new RunProperties();
            RunFonts runFonts60 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold44 = new Bold();
            BoldComplexScript boldComplexScript23 = new BoldComplexScript();
            FontSize fontSize65 = new FontSize() { Val = "28" };

            runProperties51.Append(runFonts60);
            runProperties51.Append(bold44);
            runProperties51.Append(boldComplexScript23);
            runProperties51.Append(fontSize65);
            Text text49 = new Text();
            text49.Text = "：";

            run51.Append(runProperties51);
            run51.Append(text49);

            paragraph34.Append(paragraphProperties32);
            paragraph34.Append(run46);
            paragraph34.Append(run47);
            paragraph34.Append(run48);
            paragraph34.Append(run49);
            paragraph34.Append(run50);
            paragraph34.Append(run51);

            Paragraph paragraph35 = new Paragraph() { RsidParagraphAddition = "00CF43E8", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "0014137D", ParagraphId = "3CCEB26C", TextId = "77777777" };

            ParagraphProperties paragraphProperties33 = new ParagraphProperties();
            SnapToGrid snapToGrid23 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines7 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "400", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation12 = new Indentation() { Start = "566", Hanging = "566", HangingChars = 202 };

            ParagraphMarkRunProperties paragraphMarkRunProperties30 = new ParagraphMarkRunProperties();
            Bold bold45 = new Bold();
            BoldComplexScript boldComplexScript24 = new BoldComplexScript();
            FontSize fontSize66 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties30.Append(bold45);
            paragraphMarkRunProperties30.Append(boldComplexScript24);
            paragraphMarkRunProperties30.Append(fontSize66);

            paragraphProperties33.Append(snapToGrid23);
            paragraphProperties33.Append(spacingBetweenLines7);
            paragraphProperties33.Append(indentation12);
            paragraphProperties33.Append(paragraphMarkRunProperties30);

            Run run52 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties52 = new RunProperties();
            RunFonts runFonts61 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold46 = new Bold();
            BoldComplexScript boldComplexScript25 = new BoldComplexScript();
            FontSize fontSize67 = new FontSize() { Val = "28" };

            runProperties52.Append(runFonts61);
            runProperties52.Append(bold46);
            runProperties52.Append(boldComplexScript25);
            runProperties52.Append(fontSize67);
            Text text50 = new Text();
            text50.Text = "五、本次查核重點";

            run52.Append(runProperties52);
            run52.Append(text50);

            Run run53 = new Run() { RsidRunProperties = "00723783", RsidRunAddition = "00A049BB" };

            RunProperties runProperties53 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold47 = new Bold();
            BoldComplexScript boldComplexScript26 = new BoldComplexScript();
            Color color29 = new Color() { Val = "000000" };
            FontSize fontSize68 = new FontSize() { Val = "28" };

            runProperties53.Append(runFonts62);
            runProperties53.Append(bold47);
            runProperties53.Append(boldComplexScript26);
            runProperties53.Append(color29);
            runProperties53.Append(fontSize68);
            Text text51 = new Text();
            text51.Text = "及抽樣標準";

            run53.Append(runProperties53);
            run53.Append(text51);

            Run run54 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00882386" };

            RunProperties runProperties54 = new RunProperties();
            RunFonts runFonts63 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            Bold bold48 = new Bold();
            BoldComplexScript boldComplexScript27 = new BoldComplexScript();

            runProperties54.Append(runFonts63);
            runProperties54.Append(bold48);
            runProperties54.Append(boldComplexScript27);
            Text text52 = new Text();
            text52.Text = "（";

            run54.Append(runProperties54);
            run54.Append(text52);

            Run run55 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "0099617D" };

            RunProperties runProperties55 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold49 = new Bold();
            BoldComplexScript boldComplexScript28 = new BoldComplexScript();

            runProperties55.Append(runFonts64);
            runProperties55.Append(bold49);
            runProperties55.Append(boldComplexScript28);
            Text text53 = new Text();
            text53.Text = "包含但不限於，";

            run55.Append(runProperties55);
            run55.Append(text53);

            Run run56 = new Run() { RsidRunProperties = "00051944" };

            RunProperties runProperties56 = new RunProperties();
            RunFonts runFonts65 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold50 = new Bold();
            BoldComplexScript boldComplexScript29 = new BoldComplexScript();

            runProperties56.Append(runFonts65);
            runProperties56.Append(bold50);
            runProperties56.Append(boldComplexScript29);
            Text text54 = new Text();
            text54.Text = "主管機關裁罰";

            run56.Append(runProperties56);
            run56.Append(text54);

            Run run57 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00FA0623" };

            RunProperties runProperties57 = new RunProperties();
            RunFonts runFonts66 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold51 = new Bold();
            BoldComplexScript boldComplexScript30 = new BoldComplexScript();

            runProperties57.Append(runFonts66);
            runProperties57.Append(bold51);
            runProperties57.Append(boldComplexScript30);
            Text text55 = new Text();
            text55.Text = "、";

            run57.Append(runProperties57);
            run57.Append(text55);

            Run run58 = new Run() { RsidRunProperties = "00051944" };

            RunProperties runProperties58 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold52 = new Bold();
            BoldComplexScript boldComplexScript31 = new BoldComplexScript();

            runProperties58.Append(runFonts67);
            runProperties58.Append(bold52);
            runProperties58.Append(boldComplexScript31);
            Text text56 = new Text();
            text56.Text = "重大偶發事件";

            run58.Append(runProperties58);
            run58.Append(text56);

            Run run59 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00BE42A7" };

            RunProperties runProperties59 = new RunProperties();
            RunFonts runFonts68 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold53 = new Bold();
            BoldComplexScript boldComplexScript32 = new BoldComplexScript();

            runProperties59.Append(runFonts68);
            runProperties59.Append(bold53);
            runProperties59.Append(boldComplexScript32);
            Text text57 = new Text();
            text57.Text = "、業務申辦或缺失改善經主管機關核示事項之遵循情形";

            run59.Append(runProperties59);
            run59.Append(text57);

            Run run60 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "003C3877" };

            RunProperties runProperties60 = new RunProperties();
            RunFonts runFonts69 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold54 = new Bold();
            BoldComplexScript boldComplexScript33 = new BoldComplexScript();

            runProperties60.Append(runFonts69);
            runProperties60.Append(bold54);
            runProperties60.Append(boldComplexScript33);
            Text text58 = new Text();
            text58.Text = "、結構型商品異常客訴案件";

            run60.Append(runProperties60);
            run60.Append(text58);

            Run run61 = new Run() { RsidRunProperties = "00051944", RsidRunAddition = "00882386" };

            RunProperties runProperties61 = new RunProperties();
            RunFonts runFonts70 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            Bold bold55 = new Bold();
            BoldComplexScript boldComplexScript34 = new BoldComplexScript();

            runProperties61.Append(runFonts70);
            runProperties61.Append(bold55);
            runProperties61.Append(boldComplexScript34);
            Text text59 = new Text();
            text59.Text = "）";

            run61.Append(runProperties61);
            run61.Append(text59);

            Run run62 = new Run() { RsidRunProperties = "00F03C81", RsidRunAddition = "00F03C81" };

            RunProperties runProperties62 = new RunProperties();
            RunStyle runStyle9 = new RunStyle() { Val = "a9" };
            BoldComplexScript boldComplexScript35 = new BoldComplexScript();
            FontSize fontSize69 = new FontSize() { Val = "28" };

            runProperties62.Append(runStyle9);
            runProperties62.Append(boldComplexScript35);
            runProperties62.Append(fontSize69);
            FootnoteReference footnoteReference3 = new FootnoteReference() { Id = 3 };

            run62.Append(runProperties62);
            run62.Append(footnoteReference3);

            Run run63 = new Run() { RsidRunProperties = "007129F1" };

            RunProperties runProperties63 = new RunProperties();
            RunFonts runFonts71 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold56 = new Bold();
            BoldComplexScript boldComplexScript36 = new BoldComplexScript();
            FontSize fontSize70 = new FontSize() { Val = "28" };

            runProperties63.Append(runFonts71);
            runProperties63.Append(bold56);
            runProperties63.Append(boldComplexScript36);
            runProperties63.Append(fontSize70);
            Text text60 = new Text();
            text60.Text = "：";

            run63.Append(runProperties63);
            run63.Append(text60);

            paragraph35.Append(paragraphProperties33);
            paragraph35.Append(run52);
            paragraph35.Append(run53);
            paragraph35.Append(run54);
            paragraph35.Append(run55);
            paragraph35.Append(run56);
            paragraph35.Append(run57);
            paragraph35.Append(run58);
            paragraph35.Append(run59);
            paragraph35.Append(run60);
            paragraph35.Append(run61);
            paragraph35.Append(run62);
            paragraph35.Append(run63);

            Paragraph paragraph36 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00FE4552", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "00CF43E8", ParagraphId = "609FB83E", TextId = "77777777" };

            ParagraphProperties paragraphProperties34 = new ParagraphProperties();
            SnapToGrid snapToGrid24 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines8 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "400", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation13 = new Indentation() { Start = "566", Hanging = "566", HangingChars = 202 };

            ParagraphMarkRunProperties paragraphMarkRunProperties31 = new ParagraphMarkRunProperties();
            Bold bold57 = new Bold();
            BoldComplexScript boldComplexScript37 = new BoldComplexScript();

            paragraphMarkRunProperties31.Append(bold57);
            paragraphMarkRunProperties31.Append(boldComplexScript37);

            paragraphProperties34.Append(snapToGrid24);
            paragraphProperties34.Append(spacingBetweenLines8);
            paragraphProperties34.Append(indentation13);
            paragraphProperties34.Append(paragraphMarkRunProperties31);

            Run run64 = new Run() { RsidRunProperties = "005F213D" };

            RunProperties runProperties64 = new RunProperties();
            RunFonts runFonts72 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold58 = new Bold();
            BoldComplexScript boldComplexScript38 = new BoldComplexScript();
            Color color30 = new Color() { Val = "000000" };
            FontSize fontSize71 = new FontSize() { Val = "28" };

            runProperties64.Append(runFonts72);
            runProperties64.Append(bold58);
            runProperties64.Append(boldComplexScript38);
            runProperties64.Append(color30);
            runProperties64.Append(fontSize71);
            Text text61 = new Text();
            text61.Text = "六、與海外內部稽核協";

            run64.Append(runProperties64);
            run64.Append(text61);

            Run run65 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties65 = new RunProperties();
            RunFonts runFonts73 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold59 = new Bold();
            BoldComplexScript boldComplexScript39 = new BoldComplexScript();
            FontSize fontSize72 = new FontSize() { Val = "28" };

            runProperties65.Append(runFonts73);
            runProperties65.Append(bold59);
            runProperties65.Append(boldComplexScript39);
            runProperties65.Append(fontSize72);
            Text text62 = new Text();
            text62.Text = "同查核";

            run65.Append(runProperties65);
            run65.Append(text62);

            Run run66 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties66 = new RunProperties();
            RunFonts runFonts74 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold60 = new Bold();
            BoldComplexScript boldComplexScript40 = new BoldComplexScript();

            runProperties66.Append(runFonts74);
            runProperties66.Append(bold60);
            runProperties66.Append(boldComplexScript40);
            Text text63 = new Text();
            text63.Text = "(";

            run66.Append(runProperties66);
            run66.Append(text63);

            Run run67 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties67 = new RunProperties();
            RunFonts runFonts75 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold61 = new Bold();
            BoldComplexScript boldComplexScript41 = new BoldComplexScript();

            runProperties67.Append(runFonts75);
            runProperties67.Append(bold61);
            runProperties67.Append(boldComplexScript41);
            Text text64 = new Text();
            text64.Text = "說明協同執行方式及海外內部稽核負責工作項目範圍";

            run67.Append(runProperties67);
            run67.Append(text64);

            Run run68 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties68 = new RunProperties();
            RunFonts runFonts76 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold62 = new Bold();
            BoldComplexScript boldComplexScript42 = new BoldComplexScript();

            runProperties68.Append(runFonts76);
            runProperties68.Append(bold62);
            runProperties68.Append(boldComplexScript42);
            Text text65 = new Text();
            text65.Text = ")";

            run68.Append(runProperties68);
            run68.Append(text65);

            paragraph36.Append(paragraphProperties34);
            paragraph36.Append(run64);
            paragraph36.Append(run65);
            paragraph36.Append(run66);
            paragraph36.Append(run67);
            paragraph36.Append(run68);

            Paragraph paragraph37 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00882386", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "00CF43E8", ParagraphId = "6A6C08A6", TextId = "77777777" };

            ParagraphProperties paragraphProperties35 = new ParagraphProperties();
            SnapToGrid snapToGrid25 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines9 = new SpacingBetweenLines() { Before = "360", BeforeLines = 100, Line = "400", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties32 = new ParagraphMarkRunProperties();
            Bold bold63 = new Bold();
            BoldComplexScript boldComplexScript43 = new BoldComplexScript();
            FontSize fontSize73 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties32.Append(bold63);
            paragraphMarkRunProperties32.Append(boldComplexScript43);
            paragraphMarkRunProperties32.Append(fontSize73);

            paragraphProperties35.Append(snapToGrid25);
            paragraphProperties35.Append(spacingBetweenLines9);
            paragraphProperties35.Append(paragraphMarkRunProperties32);

            Run run69 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties69 = new RunProperties();
            RunFonts runFonts77 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold64 = new Bold();
            BoldComplexScript boldComplexScript44 = new BoldComplexScript();
            FontSize fontSize74 = new FontSize() { Val = "28" };

            runProperties69.Append(runFonts77);
            runProperties69.Append(bold64);
            runProperties69.Append(boldComplexScript44);
            runProperties69.Append(fontSize74);
            LastRenderedPageBreak lastRenderedPageBreak1 = new LastRenderedPageBreak();
            Text text66 = new Text();
            text66.Text = "七";

            run69.Append(runProperties69);
            run69.Append(lastRenderedPageBreak1);
            run69.Append(text66);

            Run run70 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "0014137D" };

            RunProperties runProperties70 = new RunProperties();
            RunFonts runFonts78 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold65 = new Bold();
            BoldComplexScript boldComplexScript45 = new BoldComplexScript();
            FontSize fontSize75 = new FontSize() { Val = "28" };

            runProperties70.Append(runFonts78);
            runProperties70.Append(bold65);
            runProperties70.Append(boldComplexScript45);
            runProperties70.Append(fontSize75);
            Text text67 = new Text();
            text67.Text = "、其他說明事項";

            run70.Append(runProperties70);
            run70.Append(text67);

            Run run71 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F87EAC" };

            RunProperties runProperties71 = new RunProperties();
            RunStyle runStyle10 = new RunStyle() { Val = "a9" };
            Bold bold66 = new Bold();
            BoldComplexScript boldComplexScript46 = new BoldComplexScript();
            FontSize fontSize76 = new FontSize() { Val = "28" };
            Underline underline5 = new Underline() { Val = UnderlineValues.Single };

            runProperties71.Append(runStyle10);
            runProperties71.Append(bold66);
            runProperties71.Append(boldComplexScript46);
            runProperties71.Append(fontSize76);
            runProperties71.Append(underline5);
            FootnoteReference footnoteReference4 = new FootnoteReference() { Id = 4 };

            run71.Append(runProperties71);
            run71.Append(footnoteReference4);

            Run run72 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "0014137D" };

            RunProperties runProperties72 = new RunProperties();
            RunFonts runFonts79 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold67 = new Bold();
            BoldComplexScript boldComplexScript47 = new BoldComplexScript();
            FontSize fontSize77 = new FontSize() { Val = "28" };

            runProperties72.Append(runFonts79);
            runProperties72.Append(bold67);
            runProperties72.Append(boldComplexScript47);
            runProperties72.Append(fontSize77);
            Text text68 = new Text();
            text68.Text = "：";

            run72.Append(runProperties72);
            run72.Append(text68);

            paragraph37.Append(paragraphProperties35);
            paragraph37.Append(run69);
            paragraph37.Append(run70);
            paragraph37.Append(run71);
            paragraph37.Append(run72);

            Paragraph paragraph38 = new Paragraph() { RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "008A0512", ParagraphId = "4D94B83F", TextId = "59EA26D0" };

            ParagraphProperties paragraphProperties36 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId10 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines10 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties33 = new ParagraphMarkRunProperties();
            Bold bold68 = new Bold();
            BoldComplexScript boldComplexScript48 = new BoldComplexScript();
            FontSize fontSize78 = new FontSize() { Val = "40" };
            Underline underline6 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties33.Append(bold68);
            paragraphMarkRunProperties33.Append(boldComplexScript48);
            paragraphMarkRunProperties33.Append(fontSize78);
            paragraphMarkRunProperties33.Append(underline6);

            paragraphProperties36.Append(paragraphStyleId10);
            paragraphProperties36.Append(spacingBetweenLines10);
            paragraphProperties36.Append(paragraphMarkRunProperties33);

            Run run73 = new Run();

            RunProperties runProperties73 = new RunProperties();
            RunFonts runFonts80 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold69 = new Bold();
            BoldComplexScript boldComplexScript49 = new BoldComplexScript();
            NoProof noProof1 = new NoProof();
            FontSize fontSize79 = new FontSize() { Val = "40" };
            Underline underline7 = new Underline() { Val = UnderlineValues.Single };

            runProperties73.Append(runFonts80);
            runProperties73.Append(bold69);
            runProperties73.Append(boldComplexScript49);
            runProperties73.Append(noProof1);
            runProperties73.Append(fontSize79);
            runProperties73.Append(underline7);

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251657728U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "6FACF970", AnchorId = "777464AB" };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "381000";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "551815";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 4966335L, Cy = 2145030L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 7620L, TopEdge = 9525L, RightEdge = 7620L, BottomEdge = 7620L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1699412372U, Name = "Text Box 4" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties1 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks1 = new A.ShapeLocks() { NoChangeArrowheads = true };

            nonVisualDrawingShapeProperties1.Append(shapeLocks1);

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 4966335L, Cy = 2145030L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill1.Append(rgbColorModelHex1);

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Round };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "000000" };

            solidFill2.Append(rgbColorModelHex2);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.SystemDot };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };
            A.HeadEnd headEnd1 = new A.HeadEnd();
            A.TailEnd tailEnd1 = new A.TailEnd();

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);
            outline1.Append(headEnd1);
            outline1.Append(tailEnd1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(solidFill1);
            shapeProperties1.Append(outline1);

            Wps.TextBoxInfo2 textBoxInfo21 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent1 = new TextBoxContent();

            Paragraph paragraph39 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "79BDD758", TextId = "77777777" };

            ParagraphProperties paragraphProperties37 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines11 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties34 = new ParagraphMarkRunProperties();
            RunFonts runFonts81 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color31 = new Color() { Val = "000000" };
            FontSize fontSize80 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript8 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties34.Append(runFonts81);
            paragraphMarkRunProperties34.Append(color31);
            paragraphMarkRunProperties34.Append(fontSize80);
            paragraphMarkRunProperties34.Append(fontSizeComplexScript8);

            paragraphProperties37.Append(spacingBetweenLines11);
            paragraphProperties37.Append(paragraphMarkRunProperties34);

            Run run74 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties74 = new RunProperties();
            RunFonts runFonts82 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color32 = new Color() { Val = "000000" };
            FontSize fontSize81 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript9 = new FontSizeComplexScript() { Val = "20" };
            Languages languages7 = new Languages() { EastAsia = "zh-HK" };

            runProperties74.Append(runFonts82);
            runProperties74.Append(color32);
            runProperties74.Append(fontSize81);
            runProperties74.Append(fontSizeComplexScript9);
            runProperties74.Append(languages7);
            Text text69 = new Text();
            text69.Text = "範本";

            run74.Append(runProperties74);
            run74.Append(text69);

            Run run75 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties75 = new RunProperties();
            RunFonts runFonts83 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color33 = new Color() { Val = "000000" };
            FontSize fontSize82 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript10 = new FontSizeComplexScript() { Val = "20" };

            runProperties75.Append(runFonts83);
            runProperties75.Append(color33);
            runProperties75.Append(fontSize82);
            runProperties75.Append(fontSizeComplexScript10);
            Text text70 = new Text();
            text70.Text = "說明：（";

            run75.Append(runProperties75);
            run75.Append(text70);

            Run run76 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties76 = new RunProperties();
            RunFonts runFonts84 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color34 = new Color() { Val = "000000" };
            FontSize fontSize83 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript11 = new FontSizeComplexScript() { Val = "20" };
            Languages languages8 = new Languages() { EastAsia = "zh-HK" };

            runProperties76.Append(runFonts84);
            runProperties76.Append(color34);
            runProperties76.Append(fontSize83);
            runProperties76.Append(fontSizeComplexScript11);
            runProperties76.Append(languages8);
            Text text71 = new Text();
            text71.Text = "正式計畫請刪除本文字方塊";

            run76.Append(runProperties76);
            run76.Append(text71);

            Run run77 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties77 = new RunProperties();
            RunFonts runFonts85 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color35 = new Color() { Val = "000000" };
            FontSize fontSize84 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript12 = new FontSizeComplexScript() { Val = "20" };

            runProperties77.Append(runFonts85);
            runProperties77.Append(color35);
            runProperties77.Append(fontSize84);
            runProperties77.Append(fontSizeComplexScript12);
            Text text72 = new Text();
            text72.Text = "）";

            run77.Append(runProperties77);
            run77.Append(text72);

            paragraph39.Append(paragraphProperties37);
            paragraph39.Append(run74);
            paragraph39.Append(run75);
            paragraph39.Append(run76);
            paragraph39.Append(run77);

            Paragraph paragraph40 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "02FB464E", TextId = "77777777" };

            ParagraphProperties paragraphProperties38 = new ParagraphProperties();

            NumberingProperties numberingProperties4 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference4 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId4 = new NumberingId() { Val = 14 };

            numberingProperties4.Append(numberingLevelReference4);
            numberingProperties4.Append(numberingId4);
            SpacingBetweenLines spacingBetweenLines12 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation14 = new Indentation() { Start = "284", Hanging = "284" };

            ParagraphMarkRunProperties paragraphMarkRunProperties35 = new ParagraphMarkRunProperties();
            RunFonts runFonts86 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            FontSize fontSize85 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript13 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties35.Append(runFonts86);
            paragraphMarkRunProperties35.Append(fontSize85);
            paragraphMarkRunProperties35.Append(fontSizeComplexScript13);

            paragraphProperties38.Append(numberingProperties4);
            paragraphProperties38.Append(spacingBetweenLines12);
            paragraphProperties38.Append(indentation14);
            paragraphProperties38.Append(paragraphMarkRunProperties35);

            Run run78 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties78 = new RunProperties();
            RunFonts runFonts87 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            FontSize fontSize86 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript14 = new FontSizeComplexScript() { Val = "20" };

            runProperties78.Append(runFonts87);
            runProperties78.Append(fontSize86);
            runProperties78.Append(fontSizeComplexScript14);
            Text text73 = new Text();
            text73.Text = "依據2018年第1季QAIP會辦結果及2018.7.4部長會議結論，「外部法規之變化」應於查核計畫中揭露，揭露內容請參考註腳說明。";

            run78.Append(runProperties78);
            run78.Append(text73);

            paragraph40.Append(paragraphProperties38);
            paragraph40.Append(run78);

            Paragraph paragraph41 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "7CC4D135", TextId = "77777777" };

            ParagraphProperties paragraphProperties39 = new ParagraphProperties();

            NumberingProperties numberingProperties5 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference5 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId5 = new NumberingId() { Val = 14 };

            numberingProperties5.Append(numberingLevelReference5);
            numberingProperties5.Append(numberingId5);
            SpacingBetweenLines spacingBetweenLines13 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation15 = new Indentation() { Start = "284", Hanging = "284" };

            ParagraphMarkRunProperties paragraphMarkRunProperties36 = new ParagraphMarkRunProperties();
            RunFonts runFonts88 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color36 = new Color() { Val = "FF0000" };
            FontSize fontSize87 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript15 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties36.Append(runFonts88);
            paragraphMarkRunProperties36.Append(color36);
            paragraphMarkRunProperties36.Append(fontSize87);
            paragraphMarkRunProperties36.Append(fontSizeComplexScript15);

            paragraphProperties39.Append(numberingProperties5);
            paragraphProperties39.Append(spacingBetweenLines13);
            paragraphProperties39.Append(indentation15);
            paragraphProperties39.Append(paragraphMarkRunProperties36);

            Run run79 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties79 = new RunProperties();
            RunFonts runFonts89 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color37 = new Color() { Val = "000000" };
            FontSize fontSize88 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript16 = new FontSizeComplexScript() { Val = "20" };

            runProperties79.Append(runFonts89);
            runProperties79.Append(color37);
            runProperties79.Append(fontSize88);
            runProperties79.Append(fontSizeComplexScript16);
            Text text74 = new Text();
            text74.Text = "依據";

            run79.Append(runProperties79);
            run79.Append(text74);

            Run run80 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties80 = new RunProperties();
            RunFonts runFonts90 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color38 = new Color() { Val = "000000" };
            FontSize fontSize89 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript17 = new FontSizeComplexScript() { Val = "20" };

            runProperties80.Append(runFonts90);
            runProperties80.Append(color38);
            runProperties80.Append(fontSize89);
            runProperties80.Append(fontSizeComplexScript17);
            Text text75 = new Text();
            text75.Text = "2022.";

            run80.Append(runProperties80);
            run80.Append(text75);

            Run run81 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties81 = new RunProperties();
            RunFonts runFonts91 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color39 = new Color() { Val = "000000" };
            FontSize fontSize90 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript18 = new FontSizeComplexScript() { Val = "20" };

            runProperties81.Append(runFonts91);
            runProperties81.Append(color39);
            runProperties81.Append(fontSize90);
            runProperties81.Append(fontSizeComplexScript18);
            Text text76 = new Text();
            text76.Text = "5";

            run81.Append(runProperties81);
            run81.Append(text76);

            Run run82 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties82 = new RunProperties();
            RunFonts runFonts92 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color40 = new Color() { Val = "000000" };
            FontSize fontSize91 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript19 = new FontSizeComplexScript() { Val = "20" };

            runProperties82.Append(runFonts92);
            runProperties82.Append(color40);
            runProperties82.Append(fontSize91);
            runProperties82.Append(fontSizeComplexScript19);
            Text text77 = new Text();
            text77.Text = ".9金";

            run82.Append(runProperties82);
            run82.Append(text77);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run83 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties83 = new RunProperties();
            RunFonts runFonts93 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color41 = new Color() { Val = "000000" };
            FontSize fontSize92 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript20 = new FontSizeComplexScript() { Val = "20" };

            runProperties83.Append(runFonts93);
            runProperties83.Append(color41);
            runProperties83.Append(fontSize92);
            runProperties83.Append(fontSizeComplexScript20);
            Text text78 = new Text();
            text78.Text = "管檢";

            run83.Append(runProperties83);
            run83.Append(text78);

            Run run84 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties84 = new RunProperties();
            RunFonts runFonts94 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color42 = new Color() { Val = "000000" };
            FontSize fontSize93 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript21 = new FontSizeComplexScript() { Val = "20" };

            runProperties84.Append(runFonts94);
            runProperties84.Append(color42);
            runProperties84.Append(fontSize93);
            runProperties84.Append(fontSizeComplexScript21);
            Text text79 = new Text();
            text79.Text = "控";

            run84.Append(runProperties84);
            run84.Append(text79);

            Run run85 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties85 = new RunProperties();
            RunFonts runFonts95 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color43 = new Color() { Val = "000000" };
            FontSize fontSize94 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript22 = new FontSizeComplexScript() { Val = "20" };

            runProperties85.Append(runFonts95);
            runProperties85.Append(color43);
            runProperties85.Append(fontSize94);
            runProperties85.Append(fontSizeComplexScript22);
            Text text80 = new Text();
            text80.Text = "字";

            run85.Append(runProperties85);
            run85.Append(text80);
            ProofError proofError2 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run86 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties86 = new RunProperties();
            RunFonts runFonts96 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color44 = new Color() { Val = "000000" };
            FontSize fontSize95 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript23 = new FontSizeComplexScript() { Val = "20" };

            runProperties86.Append(runFonts96);
            runProperties86.Append(color44);
            runProperties86.Append(fontSize95);
            runProperties86.Append(fontSizeComplexScript23);
            Text text81 = new Text();
            text81.Text = "第111060";

            run86.Append(runProperties86);
            run86.Append(text81);

            Run run87 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties87 = new RunProperties();
            RunFonts runFonts97 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color45 = new Color() { Val = "000000" };
            FontSize fontSize96 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript24 = new FontSizeComplexScript() { Val = "20" };

            runProperties87.Append(runFonts97);
            runProperties87.Append(color45);
            runProperties87.Append(fontSize96);
            runProperties87.Append(fontSizeComplexScript24);
            Text text82 = new Text();
            text82.Text = "2083";

            run87.Append(runProperties87);
            run87.Append(text82);

            Run run88 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties88 = new RunProperties();
            RunFonts runFonts98 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color46 = new Color() { Val = "000000" };
            FontSize fontSize97 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript25 = new FontSizeComplexScript() { Val = "20" };

            runProperties88.Append(runFonts98);
            runProperties88.Append(color46);
            runProperties88.Append(fontSize97);
            runProperties88.Append(fontSizeComplexScript25);
            Text text83 = new Text();
            text83.Text = "號函";

            run88.Append(runProperties88);
            run88.Append(text83);

            Run run89 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties89 = new RunProperties();
            RunFonts runFonts99 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color47 = new Color() { Val = "000000" };
            FontSize fontSize98 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript26 = new FontSizeComplexScript() { Val = "20" };

            runProperties89.Append(runFonts99);
            runProperties89.Append(color47);
            runProperties89.Append(fontSize98);
            runProperties89.Append(fontSizeComplexScript26);
            Text text84 = new Text();
            text84.Text = "所列[金管會統整歸納之內部稽核作業共通性缺失]及2022.6科長會議結論：安排一般查核之行程時，業務稽核主管應注意避免連續2年於相同時間對同一受查單位辦理查核；如係因應特殊狀況，需調整原規劃查核時程後有此情形者(如：因應";

            run89.Append(runProperties89);
            run89.Append(text84);
            ProofError proofError3 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run90 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties90 = new RunProperties();
            RunFonts runFonts100 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color48 = new Color() { Val = "000000" };
            FontSize fontSize99 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript27 = new FontSizeComplexScript() { Val = "20" };

            runProperties90.Append(runFonts100);
            runProperties90.Append(color48);
            runProperties90.Append(fontSize99);
            runProperties90.Append(fontSizeComplexScript27);
            Text text85 = new Text();
            text85.Text = "疫";

            run90.Append(runProperties90);
            run90.Append(text85);
            ProofError proofError4 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run91 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties91 = new RunProperties();
            RunFonts runFonts101 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color49 = new Color() { Val = "000000" };
            FontSize fontSize100 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript28 = new FontSizeComplexScript() { Val = "20" };

            runProperties91.Append(runFonts101);
            runProperties91.Append(color49);
            runProperties91.Append(fontSize100);
            runProperties91.Append(fontSizeComplexScript28);
            Text text86 = new Text();
            text86.Text = "情影響、或與外部金檢時程重疊調整內部稽核查核時程…等情況)，應事先陳報總稽核知悉。";

            run91.Append(runProperties91);
            run91.Append(text86);

            paragraph41.Append(paragraphProperties39);
            paragraph41.Append(run79);
            paragraph41.Append(run80);
            paragraph41.Append(run81);
            paragraph41.Append(run82);
            paragraph41.Append(proofError1);
            paragraph41.Append(run83);
            paragraph41.Append(run84);
            paragraph41.Append(run85);
            paragraph41.Append(proofError2);
            paragraph41.Append(run86);
            paragraph41.Append(run87);
            paragraph41.Append(run88);
            paragraph41.Append(run89);
            paragraph41.Append(proofError3);
            paragraph41.Append(run90);
            paragraph41.Append(proofError4);
            paragraph41.Append(run91);

            Paragraph paragraph42 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "259AEB46", TextId = "77777777" };

            ParagraphProperties paragraphProperties40 = new ParagraphProperties();

            NumberingProperties numberingProperties6 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference6 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId6 = new NumberingId() { Val = 14 };

            numberingProperties6.Append(numberingLevelReference6);
            numberingProperties6.Append(numberingId6);
            SpacingBetweenLines spacingBetweenLines14 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation16 = new Indentation() { Start = "284", Hanging = "284" };

            ParagraphMarkRunProperties paragraphMarkRunProperties37 = new ParagraphMarkRunProperties();
            RunFonts runFonts102 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color50 = new Color() { Val = "000000" };
            FontSize fontSize101 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript29 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties37.Append(runFonts102);
            paragraphMarkRunProperties37.Append(color50);
            paragraphMarkRunProperties37.Append(fontSize101);
            paragraphMarkRunProperties37.Append(fontSizeComplexScript29);

            paragraphProperties40.Append(numberingProperties6);
            paragraphProperties40.Append(spacingBetweenLines14);
            paragraphProperties40.Append(indentation16);
            paragraphProperties40.Append(paragraphMarkRunProperties37);

            Run run92 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties92 = new RunProperties();
            RunFonts runFonts103 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color51 = new Color() { Val = "FF0000" };
            FontSize fontSize102 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript30 = new FontSizeComplexScript() { Val = "20" };

            runProperties92.Append(runFonts103);
            runProperties92.Append(color51);
            runProperties92.Append(fontSize102);
            runProperties92.Append(fontSizeComplexScript30);
            Text text87 = new Text();
            text87.Text = "依據銀行一般業務檢查(";

            run92.Append(runProperties92);
            run92.Append(text87);

            Run run93 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties93 = new RunProperties();
            RunFonts runFonts104 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color52 = new Color() { Val = "FF0000" };
            FontSize fontSize103 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript31 = new FontSizeComplexScript() { Val = "20" };

            runProperties93.Append(runFonts104);
            runProperties93.Append(color52);
            runProperties93.Append(fontSize103);
            runProperties93.Append(fontSizeComplexScript31);
            Text text88 = new Text();
            text88.Text = "111H027)";

            run93.Append(runProperties93);
            run93.Append(text88);

            Run run94 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties94 = new RunProperties();
            RunFonts runFonts105 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color53 = new Color() { Val = "FF0000" };
            FontSize fontSize104 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript32 = new FontSizeComplexScript() { Val = "20" };

            runProperties94.Append(runFonts105);
            runProperties94.Append(color53);
            runProperties94.Append(fontSize104);
            runProperties94.Append(fontSizeComplexScript32);
            Text text89 = new Text();
            text89.Text = "面請意見，增加就最近一次年度內部稽核風險評估結果為高風險之";

            run94.Append(runProperties94);
            run94.Append(text89);

            Run run95 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties95 = new RunProperties();
            RunFonts runFonts106 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color54 = new Color() { Val = "FF0000" };
            FontSize fontSize105 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript33 = new FontSizeComplexScript() { Val = "20" };

            runProperties95.Append(runFonts106);
            runProperties95.Append(color54);
            runProperties95.Append(fontSize105);
            runProperties95.Append(fontSizeComplexScript33);
            Text text90 = new Text();
            text90.Text = "[";

            run95.Append(runProperties95);
            run95.Append(text90);

            Run run96 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties96 = new RunProperties();
            RunFonts runFonts107 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color55 = new Color() { Val = "FF0000" };
            FontSize fontSize106 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript34 = new FontSizeComplexScript() { Val = "20" };

            runProperties96.Append(runFonts107);
            runProperties96.Append(color55);
            runProperties96.Append(fontSize106);
            runProperties96.Append(fontSizeComplexScript34);
            Text text91 = new Text();
            text91.Text = "國內分行";

            run96.Append(runProperties96);
            run96.Append(text91);

            Run run97 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties97 = new RunProperties();
            RunFonts runFonts108 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color56 = new Color() { Val = "FF0000" };
            FontSize fontSize107 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript35 = new FontSizeComplexScript() { Val = "20" };

            runProperties97.Append(runFonts108);
            runProperties97.Append(color56);
            runProperties97.Append(fontSize107);
            runProperties97.Append(fontSizeComplexScript35);
            Text text92 = new Text();
            text92.Text = "]";

            run97.Append(runProperties97);
            run97.Append(text92);

            Run run98 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties98 = new RunProperties();
            RunFonts runFonts109 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color57 = new Color() { Val = "FF0000" };
            FontSize fontSize108 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript36 = new FontSizeComplexScript() { Val = "20" };

            runProperties98.Append(runFonts109);
            runProperties98.Append(color57);
            runProperties98.Append(fontSize108);
            runProperties98.Append(fontSizeComplexScript36);
            Text text93 = new Text();
            text93.Text = "及";

            run98.Append(runProperties98);
            run98.Append(text93);

            Run run99 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties99 = new RunProperties();
            RunFonts runFonts110 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color58 = new Color() { Val = "FF0000" };
            FontSize fontSize109 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript37 = new FontSizeComplexScript() { Val = "20" };

            runProperties99.Append(runFonts110);
            runProperties99.Append(color58);
            runProperties99.Append(fontSize109);
            runProperties99.Append(fontSizeComplexScript37);
            Text text94 = new Text();
            text94.Text = "[";

            run99.Append(runProperties99);
            run99.Append(text94);
            ProofError proofError5 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run100 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties100 = new RunProperties();
            RunFonts runFonts111 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color59 = new Color() { Val = "FF0000" };
            FontSize fontSize110 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript38 = new FontSizeComplexScript() { Val = "20" };

            runProperties100.Append(runFonts111);
            runProperties100.Append(color59);
            runProperties100.Append(fontSize110);
            runProperties100.Append(fontSizeComplexScript38);
            Text text95 = new Text();
            text95.Text = "個";

            run100.Append(runProperties100);
            run100.Append(text95);
            ProofError proofError6 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run101 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties101 = new RunProperties();
            RunFonts runFonts112 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color60 = new Color() { Val = "FF0000" };
            FontSize fontSize111 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript39 = new FontSizeComplexScript() { Val = "20" };

            runProperties101.Append(runFonts112);
            runProperties101.Append(color60);
            runProperties101.Append(fontSize111);
            runProperties101.Append(fontSizeComplexScript39);
            Text text96 = new Text();
            text96.Text = "金區域中心";

            run101.Append(runProperties101);
            run101.Append(text96);

            Run run102 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties102 = new RunProperties();
            RunFonts runFonts113 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color61 = new Color() { Val = "FF0000" };
            FontSize fontSize112 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript40 = new FontSizeComplexScript() { Val = "20" };

            runProperties102.Append(runFonts113);
            runProperties102.Append(color61);
            runProperties102.Append(fontSize112);
            runProperties102.Append(fontSizeComplexScript40);
            Text text97 = new Text();
            text97.Text = "]";

            run102.Append(runProperties102);
            run102.Append(text97);

            Run run103 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties103 = new RunProperties();
            RunFonts runFonts114 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color62 = new Color() { Val = "FF0000" };
            FontSize fontSize113 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript41 = new FontSizeComplexScript() { Val = "20" };

            runProperties103.Append(runFonts114);
            runProperties103.Append(color62);
            runProperties103.Append(fontSize113);
            runProperties103.Append(fontSizeComplexScript41);
            Text text98 = new Text();
            text98.Text = "，應於「七、其他說明事項」揭露風險評估結果並辦理受查單位管理階層風險意識評估作業。";

            run103.Append(runProperties103);
            run103.Append(text98);

            paragraph42.Append(paragraphProperties40);
            paragraph42.Append(run92);
            paragraph42.Append(run93);
            paragraph42.Append(run94);
            paragraph42.Append(run95);
            paragraph42.Append(run96);
            paragraph42.Append(run97);
            paragraph42.Append(run98);
            paragraph42.Append(run99);
            paragraph42.Append(proofError5);
            paragraph42.Append(run100);
            paragraph42.Append(proofError6);
            paragraph42.Append(run101);
            paragraph42.Append(run102);
            paragraph42.Append(run103);

            textBoxContent1.Append(paragraph39);
            textBoxContent1.Append(paragraph40);
            textBoxContent1.Append(paragraph41);
            textBoxContent1.Append(paragraph42);

            textBoxInfo21.Append(textBoxContent1);

            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties() { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 91440, TopInset = 45720, RightInset = 91440, BottomInset = 45720, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, UpRight = true };
            A.NoAutoFit noAutoFit1 = new A.NoAutoFit();

            textBodyProperties1.Append(noAutoFit1);

            wordprocessingShape1.Append(nonVisualDrawingShapeProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBoxInfo21);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
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
            shapetype1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "777464AB"));
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            V.Path path1 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

            shapetype1.Append(stroke1);
            shapetype1.Append(path1);

            V.Shape shape1 = new V.Shape() { Id = "Text Box 4", Style = "position:absolute;margin-left:30pt;margin-top:43.45pt;width:391.05pt;height:168.9pt;z-index:251657728;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page;v-text-anchor:top", OptionalString = "_x0000_s1026", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQBVEQpLLAIAAFAEAAAOAAAAZHJzL2Uyb0RvYy54bWysVNtu2zAMfR+wfxD0vjpJk6416hRdsw4D\nugvQ7QMYWY6FyaJGKbGzry8lp2nRbS/D/CCIonRInkP68mrorNhpCgZdJacnEym0U1gbt6nk92+3\nb86lCBFcDRadruReB3m1fP3qsvelnmGLttYkGMSFsveVbGP0ZVEE1eoOwgl67djZIHUQ2aRNURP0\njN7ZYjaZnBU9Uu0JlQ6BT1ejUy4zftNoFb80TdBR2EpybjGvlNd1WovlJZQbAt8adUgD/iGLDozj\noEeoFUQQWzK/QXVGEQZs4onCrsCmMUrnGria6eRFNfcteJ1rYXKCP9IU/h+s+ry7919JxOEdDixg\nLiL4O1Q/gnB404Lb6Gsi7FsNNQeeJsqK3ofy8DRRHcqQQNb9J6xZZNhGzEBDQ11ihesUjM4C7I+k\n6yEKxYfzi7Oz09OFFIp9s+l8MTnNshRQPj73FOIHjZ1Im0oSq5rhYXcXYkoHyscrKVpAa+pbY202\naLO+sSR2wB1wm79cwYtr1om+kheLWUoEuBHJ1SMXfwWb5O9PYCmZFYR2DBr2YYVx7LTORG53a7pK\nnh+fQ5m4fe/q3IwRjB33XJZ1B7ITvyPTcVgPfDGRvsZ6z7QTjm3NY8ibFumXFD23dCXDzy2QlsJ+\ndCzdxXQ+TzOQjfni7YwNeu5ZP/eAUwxVySjFuL2J49xsPZlNy5HGZnF4zXI3JgvxlNUhb27brM9h\nxNJcPLfzracfwfIBAAD//wMAUEsDBBQABgAIAAAAIQBMxCzD4AAAAAkBAAAPAAAAZHJzL2Rvd25y\nZXYueG1sTI9BT4NAFITvJv6HzTPxZpcSghR5NGpiTBoPlqrnBZ4sKfsW2W3Bf+960uNkJjPfFNvF\nDOJMk+stI6xXEQjixrY9dwhvh6ebDITzils1WCaEb3KwLS8vCpW3duY9nSvfiVDCLlcI2vsxl9I1\nmoxyKzsSB+/TTkb5IKdOtpOaQ7kZZBxFqTSq57Cg1UiPmppjdTIIh2qz241p/Trr5eXr4715kMfn\nPeL11XJ/B8LT4v/C8Isf0KEMTLU9cevEgJBG4YpHyNINiOBnSbwGUSMkcXILsizk/wflDwAAAP//\nAwBQSwECLQAUAAYACAAAACEAtoM4kv4AAADhAQAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRf\nVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAAAAAAAAAAAAAAC8BAABf\ncmVscy8ucmVsc1BLAQItABQABgAIAAAAIQBVEQpLLAIAAFAEAAAOAAAAAAAAAAAAAAAAAC4CAABk\ncnMvZTJvRG9jLnhtbFBLAQItABQABgAIAAAAIQBMxCzD4AAAAAkBAAAPAAAAAAAAAAAAAAAAAIYE\nAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAQABADzAAAAkwUAAAAA\n" };
            V.Stroke stroke2 = new V.Stroke() { EndCap = V.StrokeEndCapValues.Round, DashStyle = "1 1" };

            V.TextBox textBox1 = new V.TextBox();

            TextBoxContent textBoxContent2 = new TextBoxContent();

            Paragraph paragraph43 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "79BDD758", TextId = "77777777" };

            ParagraphProperties paragraphProperties41 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines15 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };

            ParagraphMarkRunProperties paragraphMarkRunProperties38 = new ParagraphMarkRunProperties();
            RunFonts runFonts115 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color63 = new Color() { Val = "000000" };
            FontSize fontSize114 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript42 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties38.Append(runFonts115);
            paragraphMarkRunProperties38.Append(color63);
            paragraphMarkRunProperties38.Append(fontSize114);
            paragraphMarkRunProperties38.Append(fontSizeComplexScript42);

            paragraphProperties41.Append(spacingBetweenLines15);
            paragraphProperties41.Append(paragraphMarkRunProperties38);

            Run run104 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties104 = new RunProperties();
            RunFonts runFonts116 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color64 = new Color() { Val = "000000" };
            FontSize fontSize115 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript43 = new FontSizeComplexScript() { Val = "20" };
            Languages languages9 = new Languages() { EastAsia = "zh-HK" };

            runProperties104.Append(runFonts116);
            runProperties104.Append(color64);
            runProperties104.Append(fontSize115);
            runProperties104.Append(fontSizeComplexScript43);
            runProperties104.Append(languages9);
            Text text99 = new Text();
            text99.Text = "範本";

            run104.Append(runProperties104);
            run104.Append(text99);

            Run run105 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties105 = new RunProperties();
            RunFonts runFonts117 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color65 = new Color() { Val = "000000" };
            FontSize fontSize116 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript44 = new FontSizeComplexScript() { Val = "20" };

            runProperties105.Append(runFonts117);
            runProperties105.Append(color65);
            runProperties105.Append(fontSize116);
            runProperties105.Append(fontSizeComplexScript44);
            Text text100 = new Text();
            text100.Text = "說明：（";

            run105.Append(runProperties105);
            run105.Append(text100);

            Run run106 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties106 = new RunProperties();
            RunFonts runFonts118 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color66 = new Color() { Val = "000000" };
            FontSize fontSize117 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript45 = new FontSizeComplexScript() { Val = "20" };
            Languages languages10 = new Languages() { EastAsia = "zh-HK" };

            runProperties106.Append(runFonts118);
            runProperties106.Append(color66);
            runProperties106.Append(fontSize117);
            runProperties106.Append(fontSizeComplexScript45);
            runProperties106.Append(languages10);
            Text text101 = new Text();
            text101.Text = "正式計畫請刪除本文字方塊";

            run106.Append(runProperties106);
            run106.Append(text101);

            Run run107 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties107 = new RunProperties();
            RunFonts runFonts119 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color67 = new Color() { Val = "000000" };
            FontSize fontSize118 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript46 = new FontSizeComplexScript() { Val = "20" };

            runProperties107.Append(runFonts119);
            runProperties107.Append(color67);
            runProperties107.Append(fontSize118);
            runProperties107.Append(fontSizeComplexScript46);
            Text text102 = new Text();
            text102.Text = "）";

            run107.Append(runProperties107);
            run107.Append(text102);

            paragraph43.Append(paragraphProperties41);
            paragraph43.Append(run104);
            paragraph43.Append(run105);
            paragraph43.Append(run106);
            paragraph43.Append(run107);

            Paragraph paragraph44 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "02FB464E", TextId = "77777777" };

            ParagraphProperties paragraphProperties42 = new ParagraphProperties();

            NumberingProperties numberingProperties7 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference7 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId7 = new NumberingId() { Val = 14 };

            numberingProperties7.Append(numberingLevelReference7);
            numberingProperties7.Append(numberingId7);
            SpacingBetweenLines spacingBetweenLines16 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation17 = new Indentation() { Start = "284", Hanging = "284" };

            ParagraphMarkRunProperties paragraphMarkRunProperties39 = new ParagraphMarkRunProperties();
            RunFonts runFonts120 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            FontSize fontSize119 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript47 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties39.Append(runFonts120);
            paragraphMarkRunProperties39.Append(fontSize119);
            paragraphMarkRunProperties39.Append(fontSizeComplexScript47);

            paragraphProperties42.Append(numberingProperties7);
            paragraphProperties42.Append(spacingBetweenLines16);
            paragraphProperties42.Append(indentation17);
            paragraphProperties42.Append(paragraphMarkRunProperties39);

            Run run108 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties108 = new RunProperties();
            RunFonts runFonts121 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            FontSize fontSize120 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript48 = new FontSizeComplexScript() { Val = "20" };

            runProperties108.Append(runFonts121);
            runProperties108.Append(fontSize120);
            runProperties108.Append(fontSizeComplexScript48);
            Text text103 = new Text();
            text103.Text = "依據2018年第1季QAIP會辦結果及2018.7.4部長會議結論，「外部法規之變化」應於查核計畫中揭露，揭露內容請參考註腳說明。";

            run108.Append(runProperties108);
            run108.Append(text103);

            paragraph44.Append(paragraphProperties42);
            paragraph44.Append(run108);

            Paragraph paragraph45 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "7CC4D135", TextId = "77777777" };

            ParagraphProperties paragraphProperties43 = new ParagraphProperties();

            NumberingProperties numberingProperties8 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference8 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId8 = new NumberingId() { Val = 14 };

            numberingProperties8.Append(numberingLevelReference8);
            numberingProperties8.Append(numberingId8);
            SpacingBetweenLines spacingBetweenLines17 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation18 = new Indentation() { Start = "284", Hanging = "284" };

            ParagraphMarkRunProperties paragraphMarkRunProperties40 = new ParagraphMarkRunProperties();
            RunFonts runFonts122 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color68 = new Color() { Val = "FF0000" };
            FontSize fontSize121 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript49 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties40.Append(runFonts122);
            paragraphMarkRunProperties40.Append(color68);
            paragraphMarkRunProperties40.Append(fontSize121);
            paragraphMarkRunProperties40.Append(fontSizeComplexScript49);

            paragraphProperties43.Append(numberingProperties8);
            paragraphProperties43.Append(spacingBetweenLines17);
            paragraphProperties43.Append(indentation18);
            paragraphProperties43.Append(paragraphMarkRunProperties40);

            Run run109 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties109 = new RunProperties();
            RunFonts runFonts123 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color69 = new Color() { Val = "000000" };
            FontSize fontSize122 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript50 = new FontSizeComplexScript() { Val = "20" };

            runProperties109.Append(runFonts123);
            runProperties109.Append(color69);
            runProperties109.Append(fontSize122);
            runProperties109.Append(fontSizeComplexScript50);
            Text text104 = new Text();
            text104.Text = "依據";

            run109.Append(runProperties109);
            run109.Append(text104);

            Run run110 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties110 = new RunProperties();
            RunFonts runFonts124 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color70 = new Color() { Val = "000000" };
            FontSize fontSize123 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript51 = new FontSizeComplexScript() { Val = "20" };

            runProperties110.Append(runFonts124);
            runProperties110.Append(color70);
            runProperties110.Append(fontSize123);
            runProperties110.Append(fontSizeComplexScript51);
            Text text105 = new Text();
            text105.Text = "2022.";

            run110.Append(runProperties110);
            run110.Append(text105);

            Run run111 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties111 = new RunProperties();
            RunFonts runFonts125 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color71 = new Color() { Val = "000000" };
            FontSize fontSize124 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript52 = new FontSizeComplexScript() { Val = "20" };

            runProperties111.Append(runFonts125);
            runProperties111.Append(color71);
            runProperties111.Append(fontSize124);
            runProperties111.Append(fontSizeComplexScript52);
            Text text106 = new Text();
            text106.Text = "5";

            run111.Append(runProperties111);
            run111.Append(text106);

            Run run112 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties112 = new RunProperties();
            RunFonts runFonts126 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color72 = new Color() { Val = "000000" };
            FontSize fontSize125 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript53 = new FontSizeComplexScript() { Val = "20" };

            runProperties112.Append(runFonts126);
            runProperties112.Append(color72);
            runProperties112.Append(fontSize125);
            runProperties112.Append(fontSizeComplexScript53);
            Text text107 = new Text();
            text107.Text = ".9金管檢";

            run112.Append(runProperties112);
            run112.Append(text107);

            Run run113 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties113 = new RunProperties();
            RunFonts runFonts127 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color73 = new Color() { Val = "000000" };
            FontSize fontSize126 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript54 = new FontSizeComplexScript() { Val = "20" };

            runProperties113.Append(runFonts127);
            runProperties113.Append(color73);
            runProperties113.Append(fontSize126);
            runProperties113.Append(fontSizeComplexScript54);
            Text text108 = new Text();
            text108.Text = "控";

            run113.Append(runProperties113);
            run113.Append(text108);

            Run run114 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties114 = new RunProperties();
            RunFonts runFonts128 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color74 = new Color() { Val = "000000" };
            FontSize fontSize127 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript55 = new FontSizeComplexScript() { Val = "20" };

            runProperties114.Append(runFonts128);
            runProperties114.Append(color74);
            runProperties114.Append(fontSize127);
            runProperties114.Append(fontSizeComplexScript55);
            Text text109 = new Text();
            text109.Text = "字第111060";

            run114.Append(runProperties114);
            run114.Append(text109);

            Run run115 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties115 = new RunProperties();
            RunFonts runFonts129 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color75 = new Color() { Val = "000000" };
            FontSize fontSize128 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript56 = new FontSizeComplexScript() { Val = "20" };

            runProperties115.Append(runFonts129);
            runProperties115.Append(color75);
            runProperties115.Append(fontSize128);
            runProperties115.Append(fontSizeComplexScript56);
            Text text110 = new Text();
            text110.Text = "2083";

            run115.Append(runProperties115);
            run115.Append(text110);

            Run run116 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties116 = new RunProperties();
            RunFonts runFonts130 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color76 = new Color() { Val = "000000" };
            FontSize fontSize129 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript57 = new FontSizeComplexScript() { Val = "20" };

            runProperties116.Append(runFonts130);
            runProperties116.Append(color76);
            runProperties116.Append(fontSize129);
            runProperties116.Append(fontSizeComplexScript57);
            Text text111 = new Text();
            text111.Text = "號函";

            run116.Append(runProperties116);
            run116.Append(text111);

            Run run117 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties117 = new RunProperties();
            RunFonts runFonts131 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color77 = new Color() { Val = "000000" };
            FontSize fontSize130 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript58 = new FontSizeComplexScript() { Val = "20" };

            runProperties117.Append(runFonts131);
            runProperties117.Append(color77);
            runProperties117.Append(fontSize130);
            runProperties117.Append(fontSizeComplexScript58);
            Text text112 = new Text();
            text112.Text = "所列[金管會統整歸納之內部稽核作業共通性缺失]及2022.6科長會議結論：安排一般查核之行程時，業務稽核主管應注意避免連續2年於相同時間對同一受查單位辦理查核；如係因應特殊狀況，需調整原規劃查核時程後有此情形者(如：因應疫情影響、或與外部金檢時程重疊調整內部稽核查核時程…等情況)，應事先陳報總稽核知悉。";

            run117.Append(runProperties117);
            run117.Append(text112);

            paragraph45.Append(paragraphProperties43);
            paragraph45.Append(run109);
            paragraph45.Append(run110);
            paragraph45.Append(run111);
            paragraph45.Append(run112);
            paragraph45.Append(run113);
            paragraph45.Append(run114);
            paragraph45.Append(run115);
            paragraph45.Append(run116);
            paragraph45.Append(run117);

            Paragraph paragraph46 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "259AEB46", TextId = "77777777" };

            ParagraphProperties paragraphProperties44 = new ParagraphProperties();

            NumberingProperties numberingProperties9 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference9 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId9 = new NumberingId() { Val = 14 };

            numberingProperties9.Append(numberingLevelReference9);
            numberingProperties9.Append(numberingId9);
            SpacingBetweenLines spacingBetweenLines18 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation19 = new Indentation() { Start = "284", Hanging = "284" };

            ParagraphMarkRunProperties paragraphMarkRunProperties41 = new ParagraphMarkRunProperties();
            RunFonts runFonts132 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color78 = new Color() { Val = "000000" };
            FontSize fontSize131 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript59 = new FontSizeComplexScript() { Val = "20" };

            paragraphMarkRunProperties41.Append(runFonts132);
            paragraphMarkRunProperties41.Append(color78);
            paragraphMarkRunProperties41.Append(fontSize131);
            paragraphMarkRunProperties41.Append(fontSizeComplexScript59);

            paragraphProperties44.Append(numberingProperties9);
            paragraphProperties44.Append(spacingBetweenLines18);
            paragraphProperties44.Append(indentation19);
            paragraphProperties44.Append(paragraphMarkRunProperties41);

            Run run118 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties118 = new RunProperties();
            RunFonts runFonts133 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color79 = new Color() { Val = "FF0000" };
            FontSize fontSize132 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript60 = new FontSizeComplexScript() { Val = "20" };

            runProperties118.Append(runFonts133);
            runProperties118.Append(color79);
            runProperties118.Append(fontSize132);
            runProperties118.Append(fontSizeComplexScript60);
            Text text113 = new Text();
            text113.Text = "依據銀行一般業務檢查(";

            run118.Append(runProperties118);
            run118.Append(text113);

            Run run119 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties119 = new RunProperties();
            RunFonts runFonts134 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color80 = new Color() { Val = "FF0000" };
            FontSize fontSize133 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript61 = new FontSizeComplexScript() { Val = "20" };

            runProperties119.Append(runFonts134);
            runProperties119.Append(color80);
            runProperties119.Append(fontSize133);
            runProperties119.Append(fontSizeComplexScript61);
            Text text114 = new Text();
            text114.Text = "111H027)";

            run119.Append(runProperties119);
            run119.Append(text114);

            Run run120 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties120 = new RunProperties();
            RunFonts runFonts135 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color81 = new Color() { Val = "FF0000" };
            FontSize fontSize134 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript62 = new FontSizeComplexScript() { Val = "20" };

            runProperties120.Append(runFonts135);
            runProperties120.Append(color81);
            runProperties120.Append(fontSize134);
            runProperties120.Append(fontSizeComplexScript62);
            Text text115 = new Text();
            text115.Text = "面請意見，增加就最近一次年度內部稽核風險評估結果為高風險之";

            run120.Append(runProperties120);
            run120.Append(text115);

            Run run121 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties121 = new RunProperties();
            RunFonts runFonts136 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color82 = new Color() { Val = "FF0000" };
            FontSize fontSize135 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript63 = new FontSizeComplexScript() { Val = "20" };

            runProperties121.Append(runFonts136);
            runProperties121.Append(color82);
            runProperties121.Append(fontSize135);
            runProperties121.Append(fontSizeComplexScript63);
            Text text116 = new Text();
            text116.Text = "[";

            run121.Append(runProperties121);
            run121.Append(text116);

            Run run122 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties122 = new RunProperties();
            RunFonts runFonts137 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color83 = new Color() { Val = "FF0000" };
            FontSize fontSize136 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript64 = new FontSizeComplexScript() { Val = "20" };

            runProperties122.Append(runFonts137);
            runProperties122.Append(color83);
            runProperties122.Append(fontSize136);
            runProperties122.Append(fontSizeComplexScript64);
            Text text117 = new Text();
            text117.Text = "國內分行";

            run122.Append(runProperties122);
            run122.Append(text117);

            Run run123 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties123 = new RunProperties();
            RunFonts runFonts138 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color84 = new Color() { Val = "FF0000" };
            FontSize fontSize137 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript65 = new FontSizeComplexScript() { Val = "20" };

            runProperties123.Append(runFonts138);
            runProperties123.Append(color84);
            runProperties123.Append(fontSize137);
            runProperties123.Append(fontSizeComplexScript65);
            Text text118 = new Text();
            text118.Text = "]";

            run123.Append(runProperties123);
            run123.Append(text118);

            Run run124 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties124 = new RunProperties();
            RunFonts runFonts139 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color85 = new Color() { Val = "FF0000" };
            FontSize fontSize138 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript66 = new FontSizeComplexScript() { Val = "20" };

            runProperties124.Append(runFonts139);
            runProperties124.Append(color85);
            runProperties124.Append(fontSize138);
            runProperties124.Append(fontSizeComplexScript66);
            Text text119 = new Text();
            text119.Text = "及";

            run124.Append(runProperties124);
            run124.Append(text119);

            Run run125 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties125 = new RunProperties();
            RunFonts runFonts140 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color86 = new Color() { Val = "FF0000" };
            FontSize fontSize139 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript67 = new FontSizeComplexScript() { Val = "20" };

            runProperties125.Append(runFonts140);
            runProperties125.Append(color86);
            runProperties125.Append(fontSize139);
            runProperties125.Append(fontSizeComplexScript67);
            Text text120 = new Text();
            text120.Text = "[";

            run125.Append(runProperties125);
            run125.Append(text120);

            Run run126 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties126 = new RunProperties();
            RunFonts runFonts141 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color87 = new Color() { Val = "FF0000" };
            FontSize fontSize140 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript68 = new FontSizeComplexScript() { Val = "20" };

            runProperties126.Append(runFonts141);
            runProperties126.Append(color87);
            runProperties126.Append(fontSize140);
            runProperties126.Append(fontSizeComplexScript68);
            Text text121 = new Text();
            text121.Text = "個金區域中心";

            run126.Append(runProperties126);
            run126.Append(text121);

            Run run127 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties127 = new RunProperties();
            RunFonts runFonts142 = new RunFonts() { Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color88 = new Color() { Val = "FF0000" };
            FontSize fontSize141 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript69 = new FontSizeComplexScript() { Val = "20" };

            runProperties127.Append(runFonts142);
            runProperties127.Append(color88);
            runProperties127.Append(fontSize141);
            runProperties127.Append(fontSizeComplexScript69);
            Text text122 = new Text();
            text122.Text = "]";

            run127.Append(runProperties127);
            run127.Append(text122);

            Run run128 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties128 = new RunProperties();
            RunFonts runFonts143 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "新細明體", HighAnsi = "新細明體", EastAsia = "新細明體" };
            Color color89 = new Color() { Val = "FF0000" };
            FontSize fontSize142 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript70 = new FontSizeComplexScript() { Val = "20" };

            runProperties128.Append(runFonts143);
            runProperties128.Append(color89);
            runProperties128.Append(fontSize142);
            runProperties128.Append(fontSizeComplexScript70);
            Text text123 = new Text();
            text123.Text = "，應於「七、其他說明事項」揭露風險評估結果並辦理受查單位管理階層風險意識評估作業。";

            run128.Append(runProperties128);
            run128.Append(text123);

            paragraph46.Append(paragraphProperties44);
            paragraph46.Append(run118);
            paragraph46.Append(run119);
            paragraph46.Append(run120);
            paragraph46.Append(run121);
            paragraph46.Append(run122);
            paragraph46.Append(run123);
            paragraph46.Append(run124);
            paragraph46.Append(run125);
            paragraph46.Append(run126);
            paragraph46.Append(run127);
            paragraph46.Append(run128);

            textBoxContent2.Append(paragraph43);
            textBoxContent2.Append(paragraph44);
            textBoxContent2.Append(paragraph45);
            textBoxContent2.Append(paragraph46);

            textBox1.Append(textBoxContent2);

            shape1.Append(stroke2);
            shape1.Append(textBox1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            alternateContentFallback1.Append(picture1);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            run73.Append(runProperties73);
            run73.Append(alternateContent1);

            paragraph38.Append(paragraphProperties36);
            paragraph38.Append(run73);

            Paragraph paragraph47 = new Paragraph() { RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00FE4552", RsidRunAdditionDefault = "00213B89", ParagraphId = "5D6D30F6", TextId = "77777777" };

            ParagraphProperties paragraphProperties45 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId11 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines19 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            ParagraphMarkRunProperties paragraphMarkRunProperties42 = new ParagraphMarkRunProperties();
            Bold bold70 = new Bold();
            BoldComplexScript boldComplexScript50 = new BoldComplexScript();
            FontSize fontSize143 = new FontSize() { Val = "40" };
            Underline underline8 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties42.Append(bold70);
            paragraphMarkRunProperties42.Append(boldComplexScript50);
            paragraphMarkRunProperties42.Append(fontSize143);
            paragraphMarkRunProperties42.Append(underline8);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "00213B89", RsidSect = "00213B89" };
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

            paragraphProperties45.Append(paragraphStyleId11);
            paragraphProperties45.Append(spacingBetweenLines19);
            paragraphProperties45.Append(paragraphMarkRunProperties42);
            paragraphProperties45.Append(sectionProperties1);

            paragraph47.Append(paragraphProperties45);

            Paragraph paragraph48 = new Paragraph() { RsidParagraphAddition = "00200048", RsidParagraphProperties = "00200048", RsidRunAdditionDefault = "002E4BA9", ParagraphId = "40FC32AC", TextId = "77777777" };

            ParagraphProperties paragraphProperties46 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId12 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines20 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification19 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties43 = new ParagraphMarkRunProperties();
            Bold bold71 = new Bold();
            BoldComplexScript boldComplexScript51 = new BoldComplexScript();
            FontSize fontSize144 = new FontSize() { Val = "40" };
            Underline underline9 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties43.Append(bold71);
            paragraphMarkRunProperties43.Append(boldComplexScript51);
            paragraphMarkRunProperties43.Append(fontSize144);
            paragraphMarkRunProperties43.Append(underline9);

            paragraphProperties46.Append(paragraphStyleId12);
            paragraphProperties46.Append(spacingBetweenLines20);
            paragraphProperties46.Append(justification19);
            paragraphProperties46.Append(paragraphMarkRunProperties43);

            Run run129 = new Run();

            RunProperties runProperties129 = new RunProperties();
            RunFonts runFonts144 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold72 = new Bold();
            BoldComplexScript boldComplexScript52 = new BoldComplexScript();
            FontSize fontSize145 = new FontSize() { Val = "40" };
            Underline underline10 = new Underline() { Val = UnderlineValues.Single };

            runProperties129.Append(runFonts144);
            runProperties129.Append(bold72);
            runProperties129.Append(boldComplexScript52);
            runProperties129.Append(fontSize145);
            runProperties129.Append(underline10);
            LastRenderedPageBreak lastRenderedPageBreak2 = new LastRenderedPageBreak();
            Text text124 = new Text();
            text124.Text = "主題式";

            run129.Append(runProperties129);
            run129.Append(lastRenderedPageBreak2);
            run129.Append(text124);

            Run run130 = new Run() { RsidRunAddition = "00200048" };

            RunProperties runProperties130 = new RunProperties();
            Bold bold73 = new Bold();
            BoldComplexScript boldComplexScript53 = new BoldComplexScript();
            FontSize fontSize146 = new FontSize() { Val = "40" };
            Underline underline11 = new Underline() { Val = UnderlineValues.Single };

            runProperties130.Append(bold73);
            runProperties130.Append(boldComplexScript53);
            runProperties130.Append(fontSize146);
            runProperties130.Append(underline11);
            Text text125 = new Text();
            text125.Text = "查核計畫";

            run130.Append(runProperties130);
            run130.Append(text125);

            paragraph48.Append(paragraphProperties46);
            paragraph48.Append(run129);
            paragraph48.Append(run130);

            Paragraph paragraph49 = new Paragraph() { RsidParagraphAddition = "003B1BF1", RsidParagraphProperties = "00200048", RsidRunAdditionDefault = "003B1BF1", ParagraphId = "2630BA24", TextId = "77777777" };

            ParagraphProperties paragraphProperties47 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId13 = new ParagraphStyleId() { Val = "a3" };
            SpacingBetweenLines spacingBetweenLines21 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Justification justification20 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties44 = new ParagraphMarkRunProperties();
            Bold bold74 = new Bold();
            BoldComplexScript boldComplexScript54 = new BoldComplexScript();
            FontSize fontSize147 = new FontSize() { Val = "40" };
            Underline underline12 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties44.Append(bold74);
            paragraphMarkRunProperties44.Append(boldComplexScript54);
            paragraphMarkRunProperties44.Append(fontSize147);
            paragraphMarkRunProperties44.Append(underline12);

            paragraphProperties47.Append(paragraphStyleId13);
            paragraphProperties47.Append(spacingBetweenLines21);
            paragraphProperties47.Append(justification20);
            paragraphProperties47.Append(paragraphMarkRunProperties44);

            paragraph49.Append(paragraphProperties47);

            Table table3 = new Table();

            TableProperties tableProperties3 = new TableProperties();
            TableWidth tableWidth3 = new TableWidth() { Width = "9781", Type = TableWidthUnitValues.Dxa };
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
            GridColumn gridColumn6 = new GridColumn() { Width = "1791" };
            GridColumn gridColumn7 = new GridColumn() { Width = "7990" };

            tableGrid3.Append(gridColumn6);
            tableGrid3.Append(gridColumn7);

            TableRow tableRow11 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "11375B6D", TextId = "77777777" };

            TableCell tableCell24 = new TableCell();

            TableCellProperties tableCellProperties24 = new TableCellProperties();
            TableCellWidth tableCellWidth24 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties24.Append(tableCellWidth24);

            Paragraph paragraph50 = new Paragraph() { RsidParagraphMarkRevision = "00C90DDE", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "3661BFC2", TextId = "77777777" };

            ParagraphProperties paragraphProperties48 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId14 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties45 = new ParagraphMarkRunProperties();
            RunFonts runFonts145 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize148 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties45.Append(runFonts145);
            paragraphMarkRunProperties45.Append(fontSize148);

            paragraphProperties48.Append(paragraphStyleId14);
            paragraphProperties48.Append(paragraphMarkRunProperties45);

            Run run131 = new Run() { RsidRunProperties = "00C90DDE" };

            RunProperties runProperties131 = new RunProperties();
            RunFonts runFonts146 = new RunFonts() { Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize149 = new FontSize() { Val = "28" };

            runProperties131.Append(runFonts146);
            runProperties131.Append(fontSize149);
            Text text126 = new Text();
            text126.Text = "查程名稱：";

            run131.Append(runProperties131);
            run131.Append(text126);

            paragraph50.Append(paragraphProperties48);
            paragraph50.Append(run131);

            tableCell24.Append(tableCellProperties24);
            tableCell24.Append(paragraph50);

            TableCell tableCell25 = new TableCell();

            TableCellProperties tableCellProperties25 = new TableCellProperties();
            TableCellWidth tableCellWidth25 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties25.Append(tableCellWidth25);

            Paragraph paragraph51 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "5F0656A8", TextId = "1AC14C9A" };

            ParagraphProperties paragraphProperties49 = new ParagraphProperties();
            SnapToGrid snapToGrid26 = new SnapToGrid() { Val = false };
            Indentation indentation20 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification21 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties46 = new ParagraphMarkRunProperties();
            RunFonts runFonts147 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold75 = new Bold();
            Color color90 = new Color() { Val = "0000FF" };
            FontSize fontSize150 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties46.Append(runFonts147);
            paragraphMarkRunProperties46.Append(bold75);
            paragraphMarkRunProperties46.Append(color90);
            paragraphMarkRunProperties46.Append(fontSize150);

            paragraphProperties49.Append(snapToGrid26);
            paragraphProperties49.Append(indentation20);
            paragraphProperties49.Append(justification21);
            paragraphProperties49.Append(paragraphMarkRunProperties46);

            Run run132 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties132 = new RunProperties();
            RunFonts runFonts148 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold76 = new Bold();
            Color color91 = new Color() { Val = "0000FF" };
            FontSize fontSize151 = new FontSize() { Val = "28" };

            runProperties132.Append(runFonts148);
            runProperties132.Append(bold76);
            runProperties132.Append(color91);
            runProperties132.Append(fontSize151);
            Text text127 = new Text();
            text127.Text = dt.Rows[0]["planname"].ToString();

            run132.Append(runProperties132);
            run132.Append(text127);

            paragraph51.Append(paragraphProperties49);
            paragraph51.Append(run132);

            tableCell25.Append(tableCellProperties25);
            tableCell25.Append(paragraph51);

            tableRow11.Append(tableCell24);
            tableRow11.Append(tableCell25);

            TableRow tableRow12 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "729718FC", TextId = "77777777" };

            TableCell tableCell26 = new TableCell();

            TableCellProperties tableCellProperties26 = new TableCellProperties();
            TableCellWidth tableCellWidth26 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties26.Append(tableCellWidth26);

            Paragraph paragraph52 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "13BB29C2", TextId = "77777777" };

            ParagraphProperties paragraphProperties50 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId15 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties50.Append(paragraphStyleId15);

            Run run133 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties133 = new RunProperties();
            RunFonts runFonts149 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize152 = new FontSize() { Val = "28" };

            runProperties133.Append(runFonts149);
            runProperties133.Append(fontSize152);
            Text text128 = new Text();
            text128.Text = "受檢單位：";

            run133.Append(runProperties133);
            run133.Append(text128);

            paragraph52.Append(paragraphProperties50);
            paragraph52.Append(run133);

            tableCell26.Append(tableCellProperties26);
            tableCell26.Append(paragraph52);

            TableCell tableCell27 = new TableCell();

            TableCellProperties tableCellProperties27 = new TableCellProperties();
            TableCellWidth tableCellWidth27 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties27.Append(tableCellWidth27);

            Paragraph paragraph53 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "0E3DC40F", TextId = "114065E9" };

            ParagraphProperties paragraphProperties51 = new ParagraphProperties();
            SnapToGrid snapToGrid27 = new SnapToGrid() { Val = false };
            Indentation indentation21 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification22 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties47 = new ParagraphMarkRunProperties();
            RunFonts runFonts150 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold77 = new Bold();
            Color color92 = new Color() { Val = "0000FF" };
            FontSize fontSize153 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties47.Append(runFonts150);
            paragraphMarkRunProperties47.Append(bold77);
            paragraphMarkRunProperties47.Append(color92);
            paragraphMarkRunProperties47.Append(fontSize153);

            paragraphProperties51.Append(snapToGrid27);
            paragraphProperties51.Append(indentation21);
            paragraphProperties51.Append(justification22);
            paragraphProperties51.Append(paragraphMarkRunProperties47);

            Run run134 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties134 = new RunProperties();
            RunFonts runFonts151 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold78 = new Bold();
            Color color93 = new Color() { Val = "0000FF" };
            FontSize fontSize154 = new FontSize() { Val = "28" };

            runProperties134.Append(runFonts151);
            runProperties134.Append(bold78);
            runProperties134.Append(color93);
            runProperties134.Append(fontSize154);
            Text text129 = new Text();
            text129.Text = dt.Rows[0]["auditplandept"].ToString();

            run134.Append(runProperties134);
            run134.Append(text129);

            paragraph53.Append(paragraphProperties51);
            paragraph53.Append(run134);

            tableCell27.Append(tableCellProperties27);
            tableCell27.Append(paragraph53);

            tableRow12.Append(tableCell26);
            tableRow12.Append(tableCell27);

            TableRow tableRow13 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "756EE083", TextId = "77777777" };

            TableCell tableCell28 = new TableCell();

            TableCellProperties tableCellProperties28 = new TableCellProperties();
            TableCellWidth tableCellWidth28 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties28.Append(tableCellWidth28);

            Paragraph paragraph54 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "29A67BD3", TextId = "77777777" };

            ParagraphProperties paragraphProperties52 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId16 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties48 = new ParagraphMarkRunProperties();
            FontSize fontSize155 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties48.Append(fontSize155);

            paragraphProperties52.Append(paragraphStyleId16);
            paragraphProperties52.Append(paragraphMarkRunProperties48);

            Run run135 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties135 = new RunProperties();
            RunFonts runFonts152 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize156 = new FontSize() { Val = "28" };

            runProperties135.Append(runFonts152);
            runProperties135.Append(fontSize156);
            Text text130 = new Text();
            text130.Text = "查核方式：";

            run135.Append(runProperties135);
            run135.Append(text130);

            paragraph54.Append(paragraphProperties52);
            paragraph54.Append(run135);

            tableCell28.Append(tableCellProperties28);
            tableCell28.Append(paragraph54);

            TableCell tableCell29 = new TableCell();

            TableCellProperties tableCellProperties29 = new TableCellProperties();
            TableCellWidth tableCellWidth29 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties29.Append(tableCellWidth29);

            Paragraph paragraph55 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "72E9264A", TextId = "4A065926" };

            ParagraphProperties paragraphProperties53 = new ParagraphProperties();
            SnapToGrid snapToGrid28 = new SnapToGrid() { Val = false };
            Indentation indentation22 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification23 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties49 = new ParagraphMarkRunProperties();
            RunFonts runFonts153 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold79 = new Bold();
            Color color94 = new Color() { Val = "0000FF" };
            FontSize fontSize157 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties49.Append(runFonts153);
            paragraphMarkRunProperties49.Append(bold79);
            paragraphMarkRunProperties49.Append(color94);
            paragraphMarkRunProperties49.Append(fontSize157);

            paragraphProperties53.Append(snapToGrid28);
            paragraphProperties53.Append(indentation22);
            paragraphProperties53.Append(justification23);
            paragraphProperties53.Append(paragraphMarkRunProperties49);

            Run run136 = new Run() { RsidRunProperties = "00C90DDE" };

            RunProperties runProperties136 = new RunProperties();
            RunFonts runFonts154 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold80 = new Bold();
            Color color95 = new Color() { Val = "0000FF" };
            FontSize fontSize158 = new FontSize() { Val = "28" };

            runProperties136.Append(runFonts154);
            runProperties136.Append(bold80);
            runProperties136.Append(color95);
            runProperties136.Append(fontSize158);
            Text text131 = new Text();
            text131.Text = dt.Rows[0]["plantype"].ToString();

            run136.Append(runProperties136);
            run136.Append(text131);

            paragraph55.Append(paragraphProperties53);
            paragraph55.Append(run136);

            tableCell29.Append(tableCellProperties29);
            tableCell29.Append(paragraph55);

            tableRow13.Append(tableCell28);
            tableRow13.Append(tableCell29);

            TableRow tableRow14 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "30D37C50", TextId = "77777777" };

            TableCell tableCell30 = new TableCell();

            TableCellProperties tableCellProperties30 = new TableCellProperties();
            TableCellWidth tableCellWidth30 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties30.Append(tableCellWidth30);

            Paragraph paragraph56 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "39A7F6B4", TextId = "77777777" };

            ParagraphProperties paragraphProperties54 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId17 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties50 = new ParagraphMarkRunProperties();
            FontSize fontSize159 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties50.Append(fontSize159);

            paragraphProperties54.Append(paragraphStyleId17);
            paragraphProperties54.Append(paragraphMarkRunProperties50);

            Run run137 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties137 = new RunProperties();
            RunFonts runFonts155 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize160 = new FontSize() { Val = "28" };

            runProperties137.Append(runFonts155);
            runProperties137.Append(fontSize160);
            Text text132 = new Text();
            text132.Text = "查核期間：";

            run137.Append(runProperties137);
            run137.Append(text132);

            paragraph56.Append(paragraphProperties54);
            paragraph56.Append(run137);

            tableCell30.Append(tableCellProperties30);
            tableCell30.Append(paragraph56);

            TableCell tableCell31 = new TableCell();

            TableCellProperties tableCellProperties31 = new TableCellProperties();
            TableCellWidth tableCellWidth31 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties31.Append(tableCellWidth31);

            Paragraph paragraph57 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "7BB02DA1", TextId = "7670FF31" };

            ParagraphProperties paragraphProperties55 = new ParagraphProperties();
            SnapToGrid snapToGrid29 = new SnapToGrid() { Val = false };
            Indentation indentation23 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification24 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties51 = new ParagraphMarkRunProperties();
            RunFonts runFonts156 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold81 = new Bold();
            Color color96 = new Color() { Val = "0000FF" };
            FontSize fontSize161 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties51.Append(runFonts156);
            paragraphMarkRunProperties51.Append(bold81);
            paragraphMarkRunProperties51.Append(color96);
            paragraphMarkRunProperties51.Append(fontSize161);

            paragraphProperties55.Append(snapToGrid29);
            paragraphProperties55.Append(indentation23);
            paragraphProperties55.Append(justification24);
            paragraphProperties55.Append(paragraphMarkRunProperties51);

            Run run138 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties138 = new RunProperties();
            RunFonts runFonts157 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold82 = new Bold();
            Color color97 = new Color() { Val = "0000FF" };
            FontSize fontSize162 = new FontSize() { Val = "28" };

            runProperties138.Append(runFonts157);
            runProperties138.Append(bold82);
            runProperties138.Append(color97);
            runProperties138.Append(fontSize162);
            Text text133 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["startdate"].ToString(), out DateTime date2))
            {
                text133.Text = date2.ToString("yyyy-mm-dd");
            }
            else
            {
                text133.Text = "";
            }

            run138.Append(runProperties138);
            run138.Append(text133);

            Run run139 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties139 = new RunProperties();
            RunFonts runFonts158 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold83 = new Bold();
            Color color98 = new Color() { Val = "0000FF" };
            FontSize fontSize163 = new FontSize() { Val = "28" };

            runProperties139.Append(runFonts158);
            runProperties139.Append(bold83);
            runProperties139.Append(color98);
            runProperties139.Append(fontSize163);
            Text text134 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text134.Text = " ~ ";

            run139.Append(runProperties139);
            run139.Append(text134);

            Run run140 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties140 = new RunProperties();
            RunFonts runFonts159 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold84 = new Bold();
            Color color99 = new Color() { Val = "0000FF" };
            FontSize fontSize164 = new FontSize() { Val = "28" };

            runProperties140.Append(runFonts159);
            runProperties140.Append(bold84);
            runProperties140.Append(color99);
            runProperties140.Append(fontSize164);
            Text text135 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["enddate"].ToString(), out DateTime date3))
            {
                text135.Text = date3.ToString("yyyy-mm-dd");
            }
            else
            {
                text135.Text = "";
            }

            run140.Append(runProperties140);
            run140.Append(text135);

            paragraph57.Append(paragraphProperties55);
            paragraph57.Append(run138);
            paragraph57.Append(run139);
            paragraph57.Append(run140);

            tableCell31.Append(tableCellProperties31);
            tableCell31.Append(paragraph57);

            tableRow14.Append(tableCell30);
            tableRow14.Append(tableCell31);

            TableRow tableRow15 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "6A3FAEEF", TextId = "77777777" };

            TableCell tableCell32 = new TableCell();

            TableCellProperties tableCellProperties32 = new TableCellProperties();
            TableCellWidth tableCellWidth32 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties32.Append(tableCellWidth32);

            Paragraph paragraph58 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "3C54E899", TextId = "77777777" };

            ParagraphProperties paragraphProperties56 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId18 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties52 = new ParagraphMarkRunProperties();
            FontSize fontSize165 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties52.Append(fontSize165);

            paragraphProperties56.Append(paragraphStyleId18);
            paragraphProperties56.Append(paragraphMarkRunProperties52);

            Run run141 = new Run() { RsidRunProperties = "00CD2680" };

            RunProperties runProperties141 = new RunProperties();
            RunFonts runFonts160 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize166 = new FontSize() { Val = "28" };

            runProperties141.Append(runFonts160);
            runProperties141.Append(fontSize166);
            Text text136 = new Text();
            text136.Text = "查核範圍：";

            run141.Append(runProperties141);
            run141.Append(text136);

            paragraph58.Append(paragraphProperties56);
            paragraph58.Append(run141);

            tableCell32.Append(tableCellProperties32);
            tableCell32.Append(paragraph58);

            TableCell tableCell33 = new TableCell();

            TableCellProperties tableCellProperties33 = new TableCellProperties();
            TableCellWidth tableCellWidth33 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties33.Append(tableCellWidth33);

            Paragraph paragraph59 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "206BFFC3", TextId = "4B112422" };

            ParagraphProperties paragraphProperties57 = new ParagraphProperties();
            SnapToGrid snapToGrid30 = new SnapToGrid() { Val = false };
            Indentation indentation24 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification25 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties53 = new ParagraphMarkRunProperties();
            RunFonts runFonts161 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold85 = new Bold();
            Color color100 = new Color() { Val = "0000FF" };
            FontSize fontSize167 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties53.Append(runFonts161);
            paragraphMarkRunProperties53.Append(bold85);
            paragraphMarkRunProperties53.Append(color100);
            paragraphMarkRunProperties53.Append(fontSize167);

            paragraphProperties57.Append(snapToGrid30);
            paragraphProperties57.Append(indentation24);
            paragraphProperties57.Append(justification25);
            paragraphProperties57.Append(paragraphMarkRunProperties53);

            Run run142 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties142 = new RunProperties();
            RunFonts runFonts162 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold86 = new Bold();
            Color color101 = new Color() { Val = "0000FF" };
            FontSize fontSize168 = new FontSize() { Val = "28" };

            runProperties142.Append(runFonts162);
            runProperties142.Append(bold86);
            runProperties142.Append(color101);
            runProperties142.Append(fontSize168);
            Text text137 = new Text();
            //text137.Text = "查核範圍起日";
            if (DateTime.TryParse(dt.Rows[0]["ar_startdate"].ToString(), out DateTime date9))
            {
                text137.Text = date9.ToString("yyyy-mm-dd");
            }
            else
            {
                text137.Text = "";
            }

            run142.Append(runProperties142);
            run142.Append(text137);

            Run run143 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties143 = new RunProperties();
            RunFonts runFonts163 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold87 = new Bold();
            Color color102 = new Color() { Val = "0000FF" };
            FontSize fontSize169 = new FontSize() { Val = "28" };

            runProperties143.Append(runFonts163);
            runProperties143.Append(bold87);
            runProperties143.Append(color102);
            runProperties143.Append(fontSize169);
            Text text138 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text138.Text = " ~ ";

            run143.Append(runProperties143);
            run143.Append(text138);

            Run run144 = new Run() { RsidRunProperties = "00C7460E" };

            RunProperties runProperties144 = new RunProperties();
            RunFonts runFonts164 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold88 = new Bold();
            Color color103 = new Color() { Val = "0000FF" };
            FontSize fontSize170 = new FontSize() { Val = "28" };

            runProperties144.Append(runFonts164);
            runProperties144.Append(bold88);
            runProperties144.Append(color103);
            runProperties144.Append(fontSize170);
            Text text139 = new Text();
            //text139.Text = "查核範圍迄日";
            if (DateTime.TryParse(dt.Rows[0]["ar_enddate"].ToString(), out DateTime date10))
            {
                text139.Text = date10.ToString("yyyy-mm-dd");
            }
            else
            {
                text139.Text = "";
            }

            run144.Append(runProperties144);
            run144.Append(text139);

            paragraph59.Append(paragraphProperties57);
            paragraph59.Append(run142);
            paragraph59.Append(run143);
            paragraph59.Append(run144);

            tableCell33.Append(tableCellProperties33);
            tableCell33.Append(paragraph59);

            tableRow15.Append(tableCell32);
            tableRow15.Append(tableCell33);

            TableRow tableRow16 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "6CF185BF", TextId = "77777777" };

            TableCell tableCell34 = new TableCell();

            TableCellProperties tableCellProperties34 = new TableCellProperties();
            TableCellWidth tableCellWidth34 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties34.Append(tableCellWidth34);

            Paragraph paragraph60 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "2826F186", TextId = "77777777" };

            ParagraphProperties paragraphProperties58 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId19 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties54 = new ParagraphMarkRunProperties();
            FontSize fontSize171 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties54.Append(fontSize171);

            paragraphProperties58.Append(paragraphStyleId19);
            paragraphProperties58.Append(paragraphMarkRunProperties54);

            Run run145 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties145 = new RunProperties();
            RunFonts runFonts165 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize172 = new FontSize() { Val = "28" };

            runProperties145.Append(runFonts165);
            runProperties145.Append(fontSize172);
            Text text140 = new Text();
            text140.Text = "領";

            run145.Append(runProperties145);
            run145.Append(text140);

            Run run146 = new Run();

            RunProperties runProperties146 = new RunProperties();
            RunFonts runFonts166 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize173 = new FontSize() { Val = "28" };

            runProperties146.Append(runFonts166);
            runProperties146.Append(fontSize173);
            Text text141 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text141.Text = "　　";

            run146.Append(runProperties146);
            run146.Append(text141);

            Run run147 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties147 = new RunProperties();
            RunFonts runFonts167 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize174 = new FontSize() { Val = "28" };

            runProperties147.Append(runFonts167);
            runProperties147.Append(fontSize174);
            Text text142 = new Text();
            text142.Text = "隊：";

            run147.Append(runProperties147);
            run147.Append(text142);

            paragraph60.Append(paragraphProperties58);
            paragraph60.Append(run145);
            paragraph60.Append(run146);
            paragraph60.Append(run147);

            tableCell34.Append(tableCellProperties34);
            tableCell34.Append(paragraph60);

            TableCell tableCell35 = new TableCell();

            TableCellProperties tableCellProperties35 = new TableCellProperties();
            TableCellWidth tableCellWidth35 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties35.Append(tableCellWidth35);

            Paragraph paragraph61 = new Paragraph() { RsidParagraphMarkRevision = "003627CF", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "790C6E54", TextId = "13919AB7" };

            ParagraphProperties paragraphProperties59 = new ParagraphProperties();
            SnapToGrid snapToGrid31 = new SnapToGrid() { Val = false };
            Indentation indentation25 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification26 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties55 = new ParagraphMarkRunProperties();
            RunFonts runFonts168 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold89 = new Bold();
            Color color104 = new Color() { Val = "0000FF" };
            FontSize fontSize175 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties55.Append(runFonts168);
            paragraphMarkRunProperties55.Append(bold89);
            paragraphMarkRunProperties55.Append(color104);
            paragraphMarkRunProperties55.Append(fontSize175);

            paragraphProperties59.Append(snapToGrid31);
            paragraphProperties59.Append(indentation25);
            paragraphProperties59.Append(justification26);
            paragraphProperties59.Append(paragraphMarkRunProperties55);

            Run run148 = new Run();

            RunProperties runProperties148 = new RunProperties();
            RunFonts runFonts169 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold90 = new Bold();
            Color color105 = new Color() { Val = "0000FF" };
            FontSize fontSize176 = new FontSize() { Val = "28" };

            runProperties148.Append(runFonts169);
            runProperties148.Append(bold90);
            runProperties148.Append(color105);
            runProperties148.Append(fontSize176);
            Text text143 = new Text();
            text143.Text = dt.Rows[0]["leader"].ToString();

            run148.Append(runProperties148);
            run148.Append(text143);

            paragraph61.Append(paragraphProperties59);
            paragraph61.Append(run148);

            tableCell35.Append(tableCellProperties35);
            tableCell35.Append(paragraph61);

            tableRow16.Append(tableCell34);
            tableRow16.Append(tableCell35);

            TableRow tableRow17 = new TableRow() { RsidTableRowMarkRevision = "00B5775C", RsidTableRowAddition = "00D62F73", RsidTableRowProperties = "008B0C6E", ParagraphId = "72534BFC", TextId = "77777777" };

            TableCell tableCell36 = new TableCell();

            TableCellProperties tableCellProperties36 = new TableCellProperties();
            TableCellWidth tableCellWidth36 = new TableCellWidth() { Width = "1791", Type = TableWidthUnitValues.Dxa };

            tableCellProperties36.Append(tableCellWidth36);

            Paragraph paragraph62 = new Paragraph() { RsidParagraphMarkRevision = "00B5775C", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "00082F71", RsidRunAdditionDefault = "00D62F73", ParagraphId = "08A4ECB2", TextId = "77777777" };

            ParagraphProperties paragraphProperties60 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId20 = new ParagraphStyleId() { Val = "a3" };

            ParagraphMarkRunProperties paragraphMarkRunProperties56 = new ParagraphMarkRunProperties();
            FontSize fontSize177 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties56.Append(fontSize177);

            paragraphProperties60.Append(paragraphStyleId20);
            paragraphProperties60.Append(paragraphMarkRunProperties56);

            Run run149 = new Run() { RsidRunProperties = "00B5775C" };

            RunProperties runProperties149 = new RunProperties();
            RunFonts runFonts170 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize178 = new FontSize() { Val = "28" };

            runProperties149.Append(runFonts170);
            runProperties149.Append(fontSize178);
            Text text144 = new Text();
            text144.Text = "查核人員：";

            run149.Append(runProperties149);
            run149.Append(text144);

            paragraph62.Append(paragraphProperties60);
            paragraph62.Append(run149);

            tableCell36.Append(tableCellProperties36);
            tableCell36.Append(paragraph62);

            TableCell tableCell37 = new TableCell();

            TableCellProperties tableCellProperties37 = new TableCellProperties();
            TableCellWidth tableCellWidth37 = new TableCellWidth() { Width = "7990", Type = TableWidthUnitValues.Dxa };

            tableCellProperties37.Append(tableCellWidth37);

            Paragraph paragraph63 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "00D62F73", RsidParagraphProperties = "008B0C6E", RsidRunAdditionDefault = "00D62F73", ParagraphId = "3A192954", TextId = "1949E1B4" };

            ParagraphProperties paragraphProperties61 = new ParagraphProperties();
            SnapToGrid snapToGrid32 = new SnapToGrid() { Val = false };
            Indentation indentation26 = new Indentation() { Start = "698", Hanging = "698", HangingChars = 249 };
            Justification justification27 = new Justification() { Val = JustificationValues.Both };

            ParagraphMarkRunProperties paragraphMarkRunProperties57 = new ParagraphMarkRunProperties();
            RunFonts runFonts171 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold91 = new Bold();
            Color color106 = new Color() { Val = "0000FF" };
            FontSize fontSize179 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties57.Append(runFonts171);
            paragraphMarkRunProperties57.Append(bold91);
            paragraphMarkRunProperties57.Append(color106);
            paragraphMarkRunProperties57.Append(fontSize179);

            paragraphProperties61.Append(snapToGrid32);
            paragraphProperties61.Append(indentation26);
            paragraphProperties61.Append(justification27);
            paragraphProperties61.Append(paragraphMarkRunProperties57);

            Run run150 = new Run() { RsidRunProperties = "000E3880" };

            RunProperties runProperties150 = new RunProperties();
            RunFonts runFonts172 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
            Bold bold92 = new Bold();
            Color color107 = new Color() { Val = "0000FF" };
            FontSize fontSize180 = new FontSize() { Val = "28" };

            runProperties150.Append(runFonts172);
            runProperties150.Append(bold92);
            runProperties150.Append(color107);
            runProperties150.Append(fontSize180);
            Text text145 = new Text();
            text145.Text = dt.Rows[0]["Member"].ToString();

            run150.Append(runProperties150);
            run150.Append(text145);

            paragraph63.Append(paragraphProperties61);
            paragraph63.Append(run150);

            tableCell37.Append(tableCellProperties37);
            tableCell37.Append(paragraph63);

            tableRow17.Append(tableCell36);
            tableRow17.Append(tableCell37);

            table3.Append(tableProperties3);
            table3.Append(tableGrid3);
            table3.Append(tableRow11);
            table3.Append(tableRow12);
            table3.Append(tableRow13);
            table3.Append(tableRow14);
            table3.Append(tableRow15);
            table3.Append(tableRow16);
            table3.Append(tableRow17);

            Paragraph paragraph64 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00200048", RsidRunAdditionDefault = "00200048", ParagraphId = "009705E6", TextId = "77777777" };

            ParagraphProperties paragraphProperties62 = new ParagraphProperties();
            SnapToGrid snapToGrid33 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines22 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };

            paragraphProperties62.Append(snapToGrid33);
            paragraphProperties62.Append(spacingBetweenLines22);

            paragraph64.Append(paragraphProperties62);

            Paragraph paragraph65 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "3CCA0C11", TextId = "77777777" };

            ParagraphProperties paragraphProperties63 = new ParagraphProperties();

            NumberingProperties numberingProperties10 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference10 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId10 = new NumberingId() { Val = 10 };

            numberingProperties10.Append(numberingLevelReference10);
            numberingProperties10.Append(numberingId10);
            SnapToGrid snapToGrid34 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines23 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation27 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties58 = new ParagraphMarkRunProperties();
            FontSize fontSize181 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties58.Append(fontSize181);

            paragraphProperties63.Append(numberingProperties10);
            paragraphProperties63.Append(snapToGrid34);
            paragraphProperties63.Append(spacingBetweenLines23);
            paragraphProperties63.Append(indentation27);
            paragraphProperties63.Append(paragraphMarkRunProperties58);

            Run run151 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties151 = new RunProperties();
            FontSize fontSize182 = new FontSize() { Val = "28" };

            runProperties151.Append(fontSize182);
            Text text146 = new Text();
            text146.Text = "「";

            run151.Append(runProperties151);
            run151.Append(text146);

            Run run152 = new Run() { RsidRunAddition = "000B740F" };

            RunProperties runProperties152 = new RunProperties();
            RunFonts runFonts173 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color108 = new Color() { Val = "0000FF" };
            FontSize fontSize183 = new FontSize() { Val = "28" };

            runProperties152.Append(runFonts173);
            runProperties152.Append(color108);
            runProperties152.Append(fontSize183);
            Text text147 = new Text();
            text147.Text = "X";

            run152.Append(runProperties152);
            run152.Append(text147);

            Run run153 = new Run() { RsidRunAddition = "000B740F" };

            RunProperties runProperties153 = new RunProperties();
            Color color109 = new Color() { Val = "0000FF" };
            FontSize fontSize184 = new FontSize() { Val = "28" };

            runProperties153.Append(color109);
            runProperties153.Append(fontSize184);
            Text text148 = new Text();
            text148.Text = "XXXXXXXX";

            run153.Append(runProperties153);
            run153.Append(text148);

            Run run154 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties154 = new RunProperties();
            FontSize fontSize185 = new FontSize() { Val = "28" };

            runProperties154.Append(fontSize185);
            Text text149 = new Text();
            text149.Text = "」之風險及目前執行情形：";

            run154.Append(runProperties154);
            run154.Append(text149);

            paragraph65.Append(paragraphProperties63);
            paragraph65.Append(run151);
            paragraph65.Append(run152);
            paragraph65.Append(run153);
            paragraph65.Append(run154);

            Paragraph paragraph66 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "2E3DB4E8", TextId = "77777777" };

            ParagraphProperties paragraphProperties64 = new ParagraphProperties();

            NumberingProperties numberingProperties11 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference11 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId11 = new NumberingId() { Val = 11 };

            numberingProperties11.Append(numberingLevelReference11);
            numberingProperties11.Append(numberingId11);
            SnapToGrid snapToGrid35 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines24 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation28 = new Indentation() { Start = "993", Hanging = "251" };

            ParagraphMarkRunProperties paragraphMarkRunProperties59 = new ParagraphMarkRunProperties();
            FontSize fontSize186 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties59.Append(fontSize186);

            paragraphProperties64.Append(numberingProperties11);
            paragraphProperties64.Append(snapToGrid35);
            paragraphProperties64.Append(spacingBetweenLines24);
            paragraphProperties64.Append(indentation28);
            paragraphProperties64.Append(paragraphMarkRunProperties59);

            Run run155 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties155 = new RunProperties();
            FontSize fontSize187 = new FontSize() { Val = "28" };

            runProperties155.Append(fontSize187);
            Text text150 = new Text();
            text150.Text = "作業流程關鍵風險點：";

            run155.Append(runProperties155);
            run155.Append(text150);

            paragraph66.Append(paragraphProperties64);
            paragraph66.Append(run155);

            Paragraph paragraph67 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "1051E739", TextId = "77777777" };

            ParagraphProperties paragraphProperties65 = new ParagraphProperties();

            NumberingProperties numberingProperties12 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference12 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId12 = new NumberingId() { Val = 11 };

            numberingProperties12.Append(numberingLevelReference12);
            numberingProperties12.Append(numberingId12);
            SnapToGrid snapToGrid36 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines25 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation29 = new Indentation() { Start = "993", Hanging = "251" };

            ParagraphMarkRunProperties paragraphMarkRunProperties60 = new ParagraphMarkRunProperties();
            FontSize fontSize188 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties60.Append(fontSize188);

            paragraphProperties65.Append(numberingProperties12);
            paragraphProperties65.Append(snapToGrid36);
            paragraphProperties65.Append(spacingBetweenLines25);
            paragraphProperties65.Append(indentation29);
            paragraphProperties65.Append(paragraphMarkRunProperties60);

            Run run156 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties156 = new RunProperties();
            FontSize fontSize189 = new FontSize() { Val = "28" };

            runProperties156.Append(fontSize189);
            Text text151 = new Text();
            text151.Text = "現有內控程序：";

            run156.Append(runProperties156);
            run156.Append(text151);

            paragraph67.Append(paragraphProperties65);
            paragraph67.Append(run156);

            Paragraph paragraph68 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "31203F5D", TextId = "77777777" };

            ParagraphProperties paragraphProperties66 = new ParagraphProperties();

            NumberingProperties numberingProperties13 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference13 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId13 = new NumberingId() { Val = 11 };

            numberingProperties13.Append(numberingLevelReference13);
            numberingProperties13.Append(numberingId13);
            SnapToGrid snapToGrid37 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines26 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation30 = new Indentation() { Start = "993", Hanging = "251" };

            ParagraphMarkRunProperties paragraphMarkRunProperties61 = new ParagraphMarkRunProperties();
            FontSize fontSize190 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties61.Append(fontSize190);

            paragraphProperties66.Append(numberingProperties13);
            paragraphProperties66.Append(snapToGrid37);
            paragraphProperties66.Append(spacingBetweenLines26);
            paragraphProperties66.Append(indentation30);
            paragraphProperties66.Append(paragraphMarkRunProperties61);

            Run run157 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties157 = new RunProperties();
            FontSize fontSize191 = new FontSize() { Val = "28" };

            runProperties157.Append(fontSize191);
            Text text152 = new Text();
            text152.Text = "現有風險缺口（或常見重大缺失）：";

            run157.Append(runProperties157);
            run157.Append(text152);

            paragraph68.Append(paragraphProperties66);
            paragraph68.Append(run157);

            Paragraph paragraph69 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "3D2B9B77", TextId = "77777777" };

            ParagraphProperties paragraphProperties67 = new ParagraphProperties();

            NumberingProperties numberingProperties14 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference14 = new NumberingLevelReference() { Val = 1 };
            NumberingId numberingId14 = new NumberingId() { Val = 11 };

            numberingProperties14.Append(numberingLevelReference14);
            numberingProperties14.Append(numberingId14);
            SnapToGrid snapToGrid38 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines27 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation31 = new Indentation() { Start = "993", Hanging = "251" };

            ParagraphMarkRunProperties paragraphMarkRunProperties62 = new ParagraphMarkRunProperties();
            FontSize fontSize192 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties62.Append(fontSize192);

            paragraphProperties67.Append(numberingProperties14);
            paragraphProperties67.Append(snapToGrid38);
            paragraphProperties67.Append(spacingBetweenLines27);
            paragraphProperties67.Append(indentation31);
            paragraphProperties67.Append(paragraphMarkRunProperties62);

            Run run158 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties158 = new RunProperties();
            FontSize fontSize193 = new FontSize() { Val = "28" };

            runProperties158.Append(fontSize193);
            Text text153 = new Text();
            text153.Text = "現有因應措施：";

            run158.Append(runProperties158);
            run158.Append(text153);

            paragraph69.Append(paragraphProperties67);
            paragraph69.Append(run158);

            Paragraph paragraph70 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "4CF80569", TextId = "77777777" };

            ParagraphProperties paragraphProperties68 = new ParagraphProperties();

            NumberingProperties numberingProperties15 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference15 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId15 = new NumberingId() { Val = 10 };

            numberingProperties15.Append(numberingLevelReference15);
            numberingProperties15.Append(numberingId15);
            SnapToGrid snapToGrid39 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines28 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation32 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties63 = new ParagraphMarkRunProperties();
            FontSize fontSize194 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties63.Append(fontSize194);

            paragraphProperties68.Append(numberingProperties15);
            paragraphProperties68.Append(snapToGrid39);
            paragraphProperties68.Append(spacingBetweenLines28);
            paragraphProperties68.Append(indentation32);
            paragraphProperties68.Append(paragraphMarkRunProperties63);

            Run run159 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties159 = new RunProperties();
            FontSize fontSize195 = new FontSize() { Val = "28" };

            runProperties159.Append(fontSize195);
            Text text154 = new Text();
            text154.Text = "主題式查核目的：";

            run159.Append(runProperties159);
            run159.Append(text154);

            paragraph70.Append(paragraphProperties68);
            paragraph70.Append(run159);

            Paragraph paragraph71 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "61F47D89", TextId = "77777777" };

            ParagraphProperties paragraphProperties69 = new ParagraphProperties();

            NumberingProperties numberingProperties16 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference16 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId16 = new NumberingId() { Val = 10 };

            numberingProperties16.Append(numberingLevelReference16);
            numberingProperties16.Append(numberingId16);
            SnapToGrid snapToGrid40 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines29 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation33 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties64 = new ParagraphMarkRunProperties();
            FontSize fontSize196 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties64.Append(fontSize196);

            paragraphProperties69.Append(numberingProperties16);
            paragraphProperties69.Append(snapToGrid40);
            paragraphProperties69.Append(spacingBetweenLines29);
            paragraphProperties69.Append(indentation33);
            paragraphProperties69.Append(paragraphMarkRunProperties64);

            Run run160 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties160 = new RunProperties();
            FontSize fontSize197 = new FontSize() { Val = "28" };

            runProperties160.Append(fontSize197);
            Text text155 = new Text();
            text155.Text = "主題式查核重點：";

            run160.Append(runProperties160);
            run160.Append(text155);

            paragraph71.Append(paragraphProperties69);
            paragraph71.Append(run160);

            Paragraph paragraph72 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "7B80DC4A", TextId = "77777777" };

            ParagraphProperties paragraphProperties70 = new ParagraphProperties();

            NumberingProperties numberingProperties17 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference17 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId17 = new NumberingId() { Val = 10 };

            numberingProperties17.Append(numberingLevelReference17);
            numberingProperties17.Append(numberingId17);
            SnapToGrid snapToGrid41 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines30 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation34 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties65 = new ParagraphMarkRunProperties();
            FontSize fontSize198 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties65.Append(fontSize198);

            paragraphProperties70.Append(numberingProperties17);
            paragraphProperties70.Append(snapToGrid41);
            paragraphProperties70.Append(spacingBetweenLines30);
            paragraphProperties70.Append(indentation34);
            paragraphProperties70.Append(paragraphMarkRunProperties65);

            Run run161 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties161 = new RunProperties();
            FontSize fontSize199 = new FontSize() { Val = "28" };

            runProperties161.Append(fontSize199);
            Text text156 = new Text();
            text156.Text = "行程安排暨工作項目分配：";

            run161.Append(runProperties161);
            run161.Append(text156);

            paragraph72.Append(paragraphProperties70);
            paragraph72.Append(run161);

            Paragraph paragraph73 = new Paragraph() { RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "198FF5D3", TextId = "77777777" };

            ParagraphProperties paragraphProperties71 = new ParagraphProperties();

            NumberingProperties numberingProperties18 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference18 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId18 = new NumberingId() { Val = 10 };

            numberingProperties18.Append(numberingLevelReference18);
            numberingProperties18.Append(numberingId18);
            SnapToGrid snapToGrid42 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines31 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation35 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties66 = new ParagraphMarkRunProperties();
            FontSize fontSize200 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties66.Append(fontSize200);

            paragraphProperties71.Append(numberingProperties18);
            paragraphProperties71.Append(snapToGrid42);
            paragraphProperties71.Append(spacingBetweenLines31);
            paragraphProperties71.Append(indentation35);
            paragraphProperties71.Append(paragraphMarkRunProperties66);

            Run run162 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties162 = new RunProperties();
            FontSize fontSize201 = new FontSize() { Val = "28" };

            runProperties162.Append(fontSize201);
            Text text157 = new Text();
            text157.Text = "資料調閱清單：";

            run162.Append(runProperties162);
            run162.Append(text157);

            paragraph73.Append(paragraphProperties71);
            paragraph73.Append(run162);

            Paragraph paragraph74 = new Paragraph() { RsidParagraphMarkRevision = "00723783", RsidParagraphAddition = "00214C36", RsidParagraphProperties = "00E6762F", RsidRunAdditionDefault = "00214C36", ParagraphId = "42D66BB4", TextId = "77777777" };

            ParagraphProperties paragraphProperties72 = new ParagraphProperties();

            NumberingProperties numberingProperties19 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference19 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId19 = new NumberingId() { Val = 10 };

            numberingProperties19.Append(numberingLevelReference19);
            numberingProperties19.Append(numberingId19);
            SnapToGrid snapToGrid43 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines32 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation36 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties67 = new ParagraphMarkRunProperties();
            Color color110 = new Color() { Val = "000000" };
            FontSize fontSize202 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties67.Append(color110);
            paragraphMarkRunProperties67.Append(fontSize202);

            paragraphProperties72.Append(numberingProperties19);
            paragraphProperties72.Append(snapToGrid43);
            paragraphProperties72.Append(spacingBetweenLines32);
            paragraphProperties72.Append(indentation36);
            paragraphProperties72.Append(paragraphMarkRunProperties67);

            Run run163 = new Run() { RsidRunProperties = "00723783" };

            RunProperties runProperties163 = new RunProperties();
            RunFonts runFonts174 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color111 = new Color() { Val = "000000" };
            FontSize fontSize203 = new FontSize() { Val = "28" };

            runProperties163.Append(runFonts174);
            runProperties163.Append(color111);
            runProperties163.Append(fontSize203);
            Text text158 = new Text();
            text158.Text = "抽樣";

            run163.Append(runProperties163);
            run163.Append(text158);

            Run run164 = new Run() { RsidRunProperties = "00723783", RsidRunAddition = "00E6762F" };

            RunProperties runProperties164 = new RunProperties();
            RunFonts runFonts175 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color112 = new Color() { Val = "000000" };
            FontSize fontSize204 = new FontSize() { Val = "28" };

            runProperties164.Append(runFonts175);
            runProperties164.Append(color112);
            runProperties164.Append(fontSize204);
            Text text159 = new Text();
            text159.Text = "標準";

            run164.Append(runProperties164);
            run164.Append(text159);

            Run run165 = new Run() { RsidRunProperties = "00723783", RsidRunAddition = "00E6762F" };

            RunProperties runProperties165 = new RunProperties();
            RunFonts runFonts176 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color113 = new Color() { Val = "000000" };
            FontSize fontSize205 = new FontSize() { Val = "28" };

            runProperties165.Append(runFonts176);
            runProperties165.Append(color113);
            runProperties165.Append(fontSize205);
            Text text160 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text160.Text = " ";

            run165.Append(runProperties165);
            run165.Append(text160);

            Run run166 = new Run() { RsidRunProperties = "00723783" };

            RunProperties runProperties166 = new RunProperties();
            RunFonts runFonts177 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color114 = new Color() { Val = "000000" };
            FontSize fontSize206 = new FontSize() { Val = "28" };

            runProperties166.Append(runFonts177);
            runProperties166.Append(color114);
            runProperties166.Append(fontSize206);
            Text text161 = new Text();
            text161.Text = "(";

            run166.Append(runProperties166);
            run166.Append(text161);

            Run run167 = new Run() { RsidRunProperties = "00723783" };

            RunProperties runProperties167 = new RunProperties();
            RunFonts runFonts178 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color115 = new Color() { Val = "000000" };
            FontSize fontSize207 = new FontSize() { Val = "28" };

            runProperties167.Append(runFonts178);
            runProperties167.Append(color115);
            runProperties167.Append(fontSize207);
            Text text162 = new Text();
            text162.Text = "含";

            run167.Append(runProperties167);
            run167.Append(text162);

            Run run168 = new Run() { RsidRunProperties = "00723783", RsidRunAddition = "00E6762F" };

            RunProperties runProperties168 = new RunProperties();
            RunFonts runFonts179 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color116 = new Color() { Val = "000000" };
            FontSize fontSize208 = new FontSize() { Val = "28" };

            runProperties168.Append(runFonts179);
            runProperties168.Append(color116);
            runProperties168.Append(fontSize208);
            Text text163 = new Text();
            text163.Text = "抽樣方式、母體範圍及";

            run168.Append(runProperties168);
            run168.Append(text163);

            Run run169 = new Run() { RsidRunProperties = "00723783" };

            RunProperties runProperties169 = new RunProperties();
            RunFonts runFonts180 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color117 = new Color() { Val = "000000" };
            FontSize fontSize209 = new FontSize() { Val = "28" };

            runProperties169.Append(runFonts180);
            runProperties169.Append(color117);
            runProperties169.Append(fontSize209);
            Text text164 = new Text();
            text164.Text = "樣本數";

            run169.Append(runProperties169);
            run169.Append(text164);

            Run run170 = new Run() { RsidRunProperties = "00723783" };

            RunProperties runProperties170 = new RunProperties();
            RunFonts runFonts181 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color118 = new Color() { Val = "000000" };
            FontSize fontSize210 = new FontSize() { Val = "28" };

            runProperties170.Append(runFonts181);
            runProperties170.Append(color118);
            runProperties170.Append(fontSize210);
            Text text165 = new Text();
            text165.Text = ")";

            run170.Append(runProperties170);
            run170.Append(text165);

            Run run171 = new Run() { RsidRunProperties = "00723783" };

            RunProperties runProperties171 = new RunProperties();
            RunFonts runFonts182 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color119 = new Color() { Val = "000000" };
            FontSize fontSize211 = new FontSize() { Val = "28" };

            runProperties171.Append(runFonts182);
            runProperties171.Append(color119);
            runProperties171.Append(fontSize211);
            Text text166 = new Text();
            text166.Text = "：";

            run171.Append(runProperties171);
            run171.Append(text166);

            paragraph74.Append(paragraphProperties72);
            paragraph74.Append(run163);
            paragraph74.Append(run164);
            paragraph74.Append(run165);
            paragraph74.Append(run166);
            paragraph74.Append(run167);
            paragraph74.Append(run168);
            paragraph74.Append(run169);
            paragraph74.Append(run170);
            paragraph74.Append(run171);

            Paragraph paragraph75 = new Paragraph() { RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "36EEC03E", TextId = "77777777" };

            ParagraphProperties paragraphProperties73 = new ParagraphProperties();

            NumberingProperties numberingProperties20 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference20 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId20 = new NumberingId() { Val = 10 };

            numberingProperties20.Append(numberingLevelReference20);
            numberingProperties20.Append(numberingId20);
            SnapToGrid snapToGrid44 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines33 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation37 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties68 = new ParagraphMarkRunProperties();
            FontSize fontSize212 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties68.Append(fontSize212);

            paragraphProperties73.Append(numberingProperties20);
            paragraphProperties73.Append(snapToGrid44);
            paragraphProperties73.Append(spacingBetweenLines33);
            paragraphProperties73.Append(indentation37);
            paragraphProperties73.Append(paragraphMarkRunProperties68);

            Run run172 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties172 = new RunProperties();
            FontSize fontSize213 = new FontSize() { Val = "28" };

            runProperties172.Append(fontSize213);
            Text text167 = new Text();
            text167.Text = "內部作業辦法、";

            run172.Append(runProperties172);
            run172.Append(text167);
            ProofError proofError7 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run173 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties173 = new RunProperties();
            FontSize fontSize214 = new FontSize() { Val = "28" };

            runProperties173.Append(fontSize214);
            Text text168 = new Text();
            text168.Text = "外部函";

            run173.Append(runProperties173);
            run173.Append(text168);
            ProofError proofError8 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run174 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties174 = new RunProperties();
            FontSize fontSize215 = new FontSize() { Val = "28" };

            runProperties174.Append(fontSize215);
            Text text169 = new Text();
            text169.Text = "令規範：";

            run174.Append(runProperties174);
            run174.Append(text169);

            paragraph75.Append(paragraphProperties73);
            paragraph75.Append(run172);
            paragraph75.Append(proofError7);
            paragraph75.Append(run173);
            paragraph75.Append(proofError8);
            paragraph75.Append(run174);

            Paragraph paragraph76 = new Paragraph() { RsidParagraphMarkRevision = "007E3224", RsidParagraphAddition = "00200048", RsidParagraphProperties = "00214C36", RsidRunAdditionDefault = "00200048", ParagraphId = "0051DAED", TextId = "77777777" };

            ParagraphProperties paragraphProperties74 = new ParagraphProperties();

            NumberingProperties numberingProperties21 = new NumberingProperties();
            NumberingLevelReference numberingLevelReference21 = new NumberingLevelReference() { Val = 0 };
            NumberingId numberingId21 = new NumberingId() { Val = 10 };

            numberingProperties21.Append(numberingLevelReference21);
            numberingProperties21.Append(numberingId21);
            SnapToGrid snapToGrid45 = new SnapToGrid() { Val = false };
            SpacingBetweenLines spacingBetweenLines34 = new SpacingBetweenLines() { Line = "276", LineRule = LineSpacingRuleValues.Auto };
            Indentation indentation38 = new Indentation() { Start = "567", Hanging = "567" };

            ParagraphMarkRunProperties paragraphMarkRunProperties69 = new ParagraphMarkRunProperties();
            FontSize fontSize216 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties69.Append(fontSize216);

            paragraphProperties74.Append(numberingProperties21);
            paragraphProperties74.Append(snapToGrid45);
            paragraphProperties74.Append(spacingBetweenLines34);
            paragraphProperties74.Append(indentation38);
            paragraphProperties74.Append(paragraphMarkRunProperties69);

            Run run175 = new Run() { RsidRunProperties = "007E3224" };

            RunProperties runProperties175 = new RunProperties();
            FontSize fontSize217 = new FontSize() { Val = "28" };

            runProperties175.Append(fontSize217);
            Text text170 = new Text();
            text170.Text = "其他說明事項：";

            run175.Append(runProperties175);
            run175.Append(text170);

            paragraph76.Append(paragraphProperties74);
            paragraph76.Append(run175);
            Paragraph paragraph77 = new Paragraph() { RsidParagraphAddition = "00200048", RsidRunAdditionDefault = "00200048", ParagraphId = "5F4865C0", TextId = "77777777" };

            SectionProperties sectionProperties2 = new SectionProperties() { RsidR = "00200048", RsidSect = "003B1BF1" };
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
            body1.Append(paragraph28);
            body1.Append(paragraph29);
            body1.Append(paragraph30);
            body1.Append(paragraph31);
            body1.Append(paragraph32);
            body1.Append(paragraph33);
            body1.Append(paragraph34);
            body1.Append(paragraph35);
            body1.Append(paragraph36);
            body1.Append(paragraph37);
            body1.Append(paragraph38);
            body1.Append(paragraph47);
            body1.Append(paragraph48);
            body1.Append(paragraph49);
            body1.Append(table3);
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
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><ct:contentTypeSchema ct:_=\"\" ma:_=\"\" ma:contentTypeName=\"文件\" ma:contentTypeID=\"0x010100D465D9393F8B7B429C8C9CE9DCEB2AA6\" ma:contentTypeVersion=\"9\" ma:contentTypeDescription=\"建立新的文件。\" ma:contentTypeScope=\"\" ma:versionID=\"1e5f73f4dc337242fde86b74da489c82\" xmlns:ct=\"http://schemas.microsoft.com/office/2006/metadata/contentType\" xmlns:ma=\"http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes\">\r\n<xsd:schema targetNamespace=\"http://schemas.microsoft.com/office/2006/metadata/properties\" ma:root=\"true\" ma:fieldsID=\"8c42c8921edb5ff244e8d475413dc0a6\" ns2:_=\"\" ns3:_=\"\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" xmlns:ns2=\"020e0566-15c6-45f6-927c-a14f1387091e\" xmlns:ns3=\"cb5579a1-84df-45aa-8c56-deace17c375d\">\r\n<xsd:import namespace=\"020e0566-15c6-45f6-927c-a14f1387091e\"/>\r\n<xsd:import namespace=\"cb5579a1-84df-45aa-8c56-deace17c375d\"/>\r\n<xsd:element name=\"properties\">\r\n<xsd:complexType>\r\n<xsd:sequence>\r\n<xsd:element name=\"documentManagement\">\r\n<xsd:complexType>\r\n<xsd:all>\r\n<xsd:element ref=\"ns2:MediaServiceMetadata\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceFastMetadata\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceObjectDetectorVersions\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:lcf76f155ced4ddcb4097134ff3c332f\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns3:TaxCatchAll\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceOCR\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceGenerationTime\" minOccurs=\"0\"/>\r\n<xsd:element ref=\"ns2:MediaServiceEventHashCode\" minOccurs=\"0\"/>\r\n</xsd:all>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:sequence>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"020e0566-15c6-45f6-927c-a14f1387091e\" elementFormDefault=\"qualified\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:dms=\"http://schemas.microsoft.com/office/2006/documentManagement/types\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\">\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/2006/documentManagement/types\"/>\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"/>\r\n<xsd:element name=\"MediaServiceMetadata\" ma:index=\"8\" nillable=\"true\" ma:displayName=\"MediaServiceMetadata\" ma:hidden=\"true\" ma:internalName=\"MediaServiceMetadata\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceFastMetadata\" ma:index=\"9\" nillable=\"true\" ma:displayName=\"MediaServiceFastMetadata\" ma:hidden=\"true\" ma:internalName=\"MediaServiceFastMetadata\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceObjectDetectorVersions\" ma:index=\"10\" nillable=\"true\" ma:displayName=\"MediaServiceObjectDetectorVersions\" ma:hidden=\"true\" ma:indexed=\"true\" ma:internalName=\"MediaServiceObjectDetectorVersions\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"lcf76f155ced4ddcb4097134ff3c332f\" ma:index=\"12\" nillable=\"true\" ma:taxonomy=\"true\" ma:internalName=\"lcf76f155ced4ddcb4097134ff3c332f\" ma:taxonomyFieldName=\"MediaServiceImageTags\" ma:displayName=\"影像標籤\" ma:readOnly=\"false\" ma:fieldId=\"{5cf76f15-5ced-4ddc-b409-7134ff3c332f}\" ma:taxonomyMulti=\"true\" ma:sspId=\"9a4d633b-e48a-43c2-b595-a1cd44314b02\" ma:termSetId=\"09814cd3-568e-fe90-9814-8d621ff8fb84\" ma:anchorId=\"fba54fb3-c3e1-fe81-a776-ca4b69148c4d\" ma:open=\"true\" ma:isKeyword=\"false\">\r\n<xsd:complexType>\r\n<xsd:sequence>\r\n<xsd:element ref=\"pc:Terms\" minOccurs=\"0\" maxOccurs=\"1\"></xsd:element>\r\n</xsd:sequence>\r\n</xsd:complexType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceOCR\" ma:index=\"14\" nillable=\"true\" ma:displayName=\"Extracted Text\" ma:internalName=\"MediaServiceOCR\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Note\">\r\n<xsd:maxLength value=\"255\"/>\r\n</xsd:restriction>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceGenerationTime\" ma:index=\"15\" nillable=\"true\" ma:displayName=\"MediaServiceGenerationTime\" ma:hidden=\"true\" ma:internalName=\"MediaServiceGenerationTime\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n<xsd:element name=\"MediaServiceEventHashCode\" ma:index=\"16\" nillable=\"true\" ma:displayName=\"MediaServiceEventHashCode\" ma:hidden=\"true\" ma:internalName=\"MediaServiceEventHashCode\" ma:readOnly=\"true\">\r\n<xsd:simpleType>\r\n<xsd:restriction base=\"dms:Text\"/>\r\n</xsd:simpleType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"cb5579a1-84df-45aa-8c56-deace17c375d\" elementFormDefault=\"qualified\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:dms=\"http://schemas.microsoft.com/office/2006/documentManagement/types\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\">\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/2006/documentManagement/types\"/>\r\n<xsd:import namespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"/>\r\n<xsd:element name=\"TaxCatchAll\" ma:index=\"13\" nillable=\"true\" ma:displayName=\"Taxonomy Catch All Column\" ma:hidden=\"true\" ma:list=\"{c0d4e7c3-a6fa-49eb-b1fa-7b786d1ce50c}\" ma:internalName=\"TaxCatchAll\" ma:showField=\"CatchAllData\" ma:web=\"cb5579a1-84df-45aa-8c56-deace17c375d\">\r\n<xsd:complexType>\r\n<xsd:complexContent>\r\n<xsd:extension base=\"dms:MultiChoiceLookup\">\r\n<xsd:sequence>\r\n<xsd:element name=\"Value\" type=\"dms:Lookup\" maxOccurs=\"unbounded\" minOccurs=\"0\" nillable=\"true\"/>\r\n</xsd:sequence>\r\n</xsd:extension>\r\n</xsd:complexContent>\r\n</xsd:complexType>\r\n</xsd:element>\r\n</xsd:schema>\r\n<xsd:schema targetNamespace=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" elementFormDefault=\"qualified\" attributeFormDefault=\"unqualified\" blockDefault=\"#all\" xmlns=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:odoc=\"http://schemas.microsoft.com/internal/obd\">\r\n<xsd:import namespace=\"http://purl.org/dc/elements/1.1/\" schemaLocation=\"http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd\"/>\r\n<xsd:import namespace=\"http://purl.org/dc/terms/\" schemaLocation=\"http://dublincore.org/schemas/xmls/qdc/2003/04/02/dcterms.xsd\"/>\r\n<xsd:element name=\"coreProperties\" type=\"CT_coreProperties\"/>\r\n<xsd:complexType name=\"CT_coreProperties\">\r\n<xsd:all>\r\n<xsd:element ref=\"dc:creator\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dcterms:created\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dc:identifier\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"contentType\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\" ma:index=\"0\" ma:displayName=\"內容類型\"/>\r\n<xsd:element ref=\"dc:title\" minOccurs=\"0\" maxOccurs=\"1\" ma:index=\"4\" ma:displayName=\"標題\"/>\r\n<xsd:element ref=\"dc:subject\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element ref=\"dc:description\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"keywords\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element ref=\"dc:language\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"category\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element name=\"version\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element name=\"revision\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\">\r\n<xsd:annotation>\r\n<xsd:documentation>\r\n                        This value indicates the number of saves or revisions. The application is responsible for updating this value after each revision.\r\n                    </xsd:documentation>\r\n</xsd:annotation>\r\n</xsd:element>\r\n<xsd:element name=\"lastModifiedBy\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n<xsd:element ref=\"dcterms:modified\" minOccurs=\"0\" maxOccurs=\"1\"/>\r\n<xsd:element name=\"contentStatus\" minOccurs=\"0\" maxOccurs=\"1\" type=\"xsd:string\"/>\r\n</xsd:all>\r\n</xsd:complexType>\r\n</xsd:schema>\r\n<xs:schema targetNamespace=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\" elementFormDefault=\"qualified\" attributeFormDefault=\"unqualified\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\">\r\n<xs:element name=\"Person\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:DisplayName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:AccountId\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:AccountType\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"DisplayName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"AccountId\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"AccountType\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"BDCAssociatedEntity\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:BDCEntity\" minOccurs=\"0\" maxOccurs=\"unbounded\"></xs:element>\r\n</xs:sequence>\r\n<xs:attribute ref=\"pc:EntityNamespace\"></xs:attribute>\r\n<xs:attribute ref=\"pc:EntityName\"></xs:attribute>\r\n<xs:attribute ref=\"pc:SystemInstanceName\"></xs:attribute>\r\n<xs:attribute ref=\"pc:AssociationName\"></xs:attribute>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:attribute name=\"EntityNamespace\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"EntityName\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"SystemInstanceName\" type=\"xs:string\"></xs:attribute>\r\n<xs:attribute name=\"AssociationName\" type=\"xs:string\"></xs:attribute>\r\n<xs:element name=\"BDCEntity\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:EntityDisplayName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityInstanceReference\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId1\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId2\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId3\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId4\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:EntityId5\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"EntityDisplayName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityInstanceReference\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId1\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId2\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId3\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId4\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"EntityId5\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"Terms\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:TermInfo\" minOccurs=\"0\" maxOccurs=\"unbounded\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"TermInfo\">\r\n<xs:complexType>\r\n<xs:sequence>\r\n<xs:element ref=\"pc:TermName\" minOccurs=\"0\"></xs:element>\r\n<xs:element ref=\"pc:TermId\" minOccurs=\"0\"></xs:element>\r\n</xs:sequence>\r\n</xs:complexType>\r\n</xs:element>\r\n<xs:element name=\"TermName\" type=\"xs:string\"></xs:element>\r\n<xs:element name=\"TermId\" type=\"xs:string\"></xs:element>\r\n</xs:schema>\r\n</ct:contentTypeSchema>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart1.
        private void GenerateCustomXmlPropertiesPart1Content(CustomXmlPropertiesPart customXmlPropertiesPart1)
        {
            Ds.DataStoreItem dataStoreItem1 = new Ds.DataStoreItem() { ItemId = "{0015AB91-9145-4166-9F51-1624D26D06EE}" };
            dataStoreItem1.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences1 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference1 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/contentType" };
            Ds.SchemaReference schemaReference2 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties/metaAttributes" };
            Ds.SchemaReference schemaReference3 = new Ds.SchemaReference() { Uri = "http://www.w3.org/2001/XMLSchema" };
            Ds.SchemaReference schemaReference4 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties" };
            Ds.SchemaReference schemaReference5 = new Ds.SchemaReference() { Uri = "020e0566-15c6-45f6-927c-a14f1387091e" };
            Ds.SchemaReference schemaReference6 = new Ds.SchemaReference() { Uri = "cb5579a1-84df-45aa-8c56-deace17c375d" };
            Ds.SchemaReference schemaReference7 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/documentManagement/types" };
            Ds.SchemaReference schemaReference8 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls" };
            Ds.SchemaReference schemaReference9 = new Ds.SchemaReference() { Uri = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" };
            Ds.SchemaReference schemaReference10 = new Ds.SchemaReference() { Uri = "http://purl.org/dc/elements/1.1/" };
            Ds.SchemaReference schemaReference11 = new Ds.SchemaReference() { Uri = "http://purl.org/dc/terms/" };
            Ds.SchemaReference schemaReference12 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/internal/obd" };

            schemaReferences1.Append(schemaReference1);
            schemaReferences1.Append(schemaReference2);
            schemaReferences1.Append(schemaReference3);
            schemaReferences1.Append(schemaReference4);
            schemaReferences1.Append(schemaReference5);
            schemaReferences1.Append(schemaReference6);
            schemaReferences1.Append(schemaReference7);
            schemaReferences1.Append(schemaReference8);
            schemaReferences1.Append(schemaReference9);
            schemaReferences1.Append(schemaReference10);
            schemaReferences1.Append(schemaReference11);
            schemaReferences1.Append(schemaReference12);

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
            CompatibilitySetting compatibilitySetting6 = new CompatibilitySetting() { Name = new EnumValue<CompatSettingNameValues>() { InnerText = "useWord2013TrackBottomHyphenation" }, Uri = "http://schemas.microsoft.com/office/word", Val = "0" };

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
            Rsid rsid2 = new Rsid() { Val = "00023ED0" };
            Rsid rsid3 = new Rsid() { Val = "00033A39" };
            Rsid rsid4 = new Rsid() { Val = "00042711" };
            Rsid rsid5 = new Rsid() { Val = "00051944" };
            Rsid rsid6 = new Rsid() { Val = "00060A25" };
            Rsid rsid7 = new Rsid() { Val = "00066F3A" };
            Rsid rsid8 = new Rsid() { Val = "00071B7E" };
            Rsid rsid9 = new Rsid() { Val = "00082F71" };
            Rsid rsid10 = new Rsid() { Val = "000A741B" };
            Rsid rsid11 = new Rsid() { Val = "000B740F" };
            Rsid rsid12 = new Rsid() { Val = "000C4189" };
            Rsid rsid13 = new Rsid() { Val = "000E0C5E" };
            Rsid rsid14 = new Rsid() { Val = "000E3880" };
            Rsid rsid15 = new Rsid() { Val = "000F43AE" };
            Rsid rsid16 = new Rsid() { Val = "00112CA3" };
            Rsid rsid17 = new Rsid() { Val = "0014137D" };
            Rsid rsid18 = new Rsid() { Val = "00143E14" };
            Rsid rsid19 = new Rsid() { Val = "001557E5" };
            Rsid rsid20 = new Rsid() { Val = "00157646" };
            Rsid rsid21 = new Rsid() { Val = "00193FEF" };
            Rsid rsid22 = new Rsid() { Val = "001A70EC" };
            Rsid rsid23 = new Rsid() { Val = "001B4FF4" };
            Rsid rsid24 = new Rsid() { Val = "001C3E6D" };
            Rsid rsid25 = new Rsid() { Val = "00200048" };
            Rsid rsid26 = new Rsid() { Val = "00210C81" };
            Rsid rsid27 = new Rsid() { Val = "00213B89" };
            Rsid rsid28 = new Rsid() { Val = "00214C36" };
            Rsid rsid29 = new Rsid() { Val = "00215F4D" };
            Rsid rsid30 = new Rsid() { Val = "00222D0C" };
            Rsid rsid31 = new Rsid() { Val = "00241F56" };
            Rsid rsid32 = new Rsid() { Val = "0028281B" };
            Rsid rsid33 = new Rsid() { Val = "002856C4" };
            Rsid rsid34 = new Rsid() { Val = "002B63CC" };
            Rsid rsid35 = new Rsid() { Val = "002E4BA9" };
            Rsid rsid36 = new Rsid() { Val = "003627CF" };
            Rsid rsid37 = new Rsid() { Val = "0036373C" };
            Rsid rsid38 = new Rsid() { Val = "00366DF0" };
            Rsid rsid39 = new Rsid() { Val = "00373FDA" };
            Rsid rsid40 = new Rsid() { Val = "003B03C7" };
            Rsid rsid41 = new Rsid() { Val = "003B1BF1" };
            Rsid rsid42 = new Rsid() { Val = "003C3877" };
            Rsid rsid43 = new Rsid() { Val = "003F2222" };
            Rsid rsid44 = new Rsid() { Val = "0040055A" };
            Rsid rsid45 = new Rsid() { Val = "00405678" };
            Rsid rsid46 = new Rsid() { Val = "00434830" };
            Rsid rsid47 = new Rsid() { Val = "004359B7" };
            Rsid rsid48 = new Rsid() { Val = "004506C8" };
            Rsid rsid49 = new Rsid() { Val = "00475D28" };
            Rsid rsid50 = new Rsid() { Val = "004869E9" };
            Rsid rsid51 = new Rsid() { Val = "004920B3" };
            Rsid rsid52 = new Rsid() { Val = "004C2CFC" };
            Rsid rsid53 = new Rsid() { Val = "0052521E" };
            Rsid rsid54 = new Rsid() { Val = "0057504C" };
            Rsid rsid55 = new Rsid() { Val = "00575F07" };
            Rsid rsid56 = new Rsid() { Val = "00585A36" };
            Rsid rsid57 = new Rsid() { Val = "00595527" };
            Rsid rsid58 = new Rsid() { Val = "005C14F0" };
            Rsid rsid59 = new Rsid() { Val = "005D39C8" };
            Rsid rsid60 = new Rsid() { Val = "005F213D" };
            Rsid rsid61 = new Rsid() { Val = "006336E9" };
            Rsid rsid62 = new Rsid() { Val = "00646300" };
            Rsid rsid63 = new Rsid() { Val = "00664C80" };
            Rsid rsid64 = new Rsid() { Val = "00676306" };
            Rsid rsid65 = new Rsid() { Val = "00686778" };
            Rsid rsid66 = new Rsid() { Val = "006C3779" };
            Rsid rsid67 = new Rsid() { Val = "006D2C24" };
            Rsid rsid68 = new Rsid() { Val = "006F6412" };
            Rsid rsid69 = new Rsid() { Val = "00704D62" };
            Rsid rsid70 = new Rsid() { Val = "007129F1" };
            Rsid rsid71 = new Rsid() { Val = "00723783" };
            Rsid rsid72 = new Rsid() { Val = "007729DD" };
            Rsid rsid73 = new Rsid() { Val = "0079412A" };
            Rsid rsid74 = new Rsid() { Val = "007D1D9B" };
            Rsid rsid75 = new Rsid() { Val = "007E23AE" };
            Rsid rsid76 = new Rsid() { Val = "00801F2D" };
            Rsid rsid77 = new Rsid() { Val = "008625F1" };
            Rsid rsid78 = new Rsid() { Val = "00866C86" };
            Rsid rsid79 = new Rsid() { Val = "00880D84" };
            Rsid rsid80 = new Rsid() { Val = "00882386" };
            Rsid rsid81 = new Rsid() { Val = "0089312B" };
            Rsid rsid82 = new Rsid() { Val = "008A0512" };
            Rsid rsid83 = new Rsid() { Val = "008A1C1B" };
            Rsid rsid84 = new Rsid() { Val = "008B0C6E" };
            Rsid rsid85 = new Rsid() { Val = "00904F8B" };
            Rsid rsid86 = new Rsid() { Val = "00920A29" };
            Rsid rsid87 = new Rsid() { Val = "009360DF" };
            Rsid rsid88 = new Rsid() { Val = "00972E50" };
            Rsid rsid89 = new Rsid() { Val = "0099123D" };
            Rsid rsid90 = new Rsid() { Val = "0099617D" };
            Rsid rsid91 = new Rsid() { Val = "009A6EFB" };
            Rsid rsid92 = new Rsid() { Val = "00A049BB" };
            Rsid rsid93 = new Rsid() { Val = "00A352E0" };
            Rsid rsid94 = new Rsid() { Val = "00A41DB3" };
            Rsid rsid95 = new Rsid() { Val = "00A46B0D" };
            Rsid rsid96 = new Rsid() { Val = "00A663F6" };
            Rsid rsid97 = new Rsid() { Val = "00A72D01" };
            Rsid rsid98 = new Rsid() { Val = "00A94093" };
            Rsid rsid99 = new Rsid() { Val = "00A96EFD" };
            Rsid rsid100 = new Rsid() { Val = "00AA4392" };
            Rsid rsid101 = new Rsid() { Val = "00AD2AF1" };
            Rsid rsid102 = new Rsid() { Val = "00AD3A9D" };
            Rsid rsid103 = new Rsid() { Val = "00AF1B88" };
            Rsid rsid104 = new Rsid() { Val = "00B4042C" };
            Rsid rsid105 = new Rsid() { Val = "00B42E31" };
            Rsid rsid106 = new Rsid() { Val = "00B65A03" };
            Rsid rsid107 = new Rsid() { Val = "00BB5086" };
            Rsid rsid108 = new Rsid() { Val = "00BE42A7" };
            Rsid rsid109 = new Rsid() { Val = "00C245D2" };
            Rsid rsid110 = new Rsid() { Val = "00C412A4" };
            Rsid rsid111 = new Rsid() { Val = "00C62CB5" };
            Rsid rsid112 = new Rsid() { Val = "00C7460E" };
            Rsid rsid113 = new Rsid() { Val = "00C75D1B" };
            Rsid rsid114 = new Rsid() { Val = "00C86D49" };
            Rsid rsid115 = new Rsid() { Val = "00C90DDE" };
            Rsid rsid116 = new Rsid() { Val = "00CA1F96" };
            Rsid rsid117 = new Rsid() { Val = "00CA7B0D" };
            Rsid rsid118 = new Rsid() { Val = "00CB0C7F" };
            Rsid rsid119 = new Rsid() { Val = "00CD2680" };
            Rsid rsid120 = new Rsid() { Val = "00CE480A" };
            Rsid rsid121 = new Rsid() { Val = "00CF43E8" };
            Rsid rsid122 = new Rsid() { Val = "00D24703" };
            Rsid rsid123 = new Rsid() { Val = "00D43514" };
            Rsid rsid124 = new Rsid() { Val = "00D62F73" };
            Rsid rsid125 = new Rsid() { Val = "00D73633" };
            Rsid rsid126 = new Rsid() { Val = "00D7709C" };
            Rsid rsid127 = new Rsid() { Val = "00D9623A" };
            Rsid rsid128 = new Rsid() { Val = "00DB5532" };
            Rsid rsid129 = new Rsid() { Val = "00DC71B6" };
            Rsid rsid130 = new Rsid() { Val = "00E2524D" };
            Rsid rsid131 = new Rsid() { Val = "00E35208" };
            Rsid rsid132 = new Rsid() { Val = "00E519A7" };
            Rsid rsid133 = new Rsid() { Val = "00E55A94" };
            Rsid rsid134 = new Rsid() { Val = "00E6762F" };
            Rsid rsid135 = new Rsid() { Val = "00EB5D9F" };
            Rsid rsid136 = new Rsid() { Val = "00ED6DFF" };
            Rsid rsid137 = new Rsid() { Val = "00ED7D88" };
            Rsid rsid138 = new Rsid() { Val = "00EE6462" };
            Rsid rsid139 = new Rsid() { Val = "00F03C81" };
            Rsid rsid140 = new Rsid() { Val = "00F45CFE" };
            Rsid rsid141 = new Rsid() { Val = "00F469D5" };
            Rsid rsid142 = new Rsid() { Val = "00F55A18" };
            Rsid rsid143 = new Rsid() { Val = "00F57CEA" };
            Rsid rsid144 = new Rsid() { Val = "00F64F71" };
            Rsid rsid145 = new Rsid() { Val = "00F87EAC" };
            Rsid rsid146 = new Rsid() { Val = "00FA0623" };
            Rsid rsid147 = new Rsid() { Val = "00FD4180" };
            Rsid rsid148 = new Rsid() { Val = "00FD740C" };
            Rsid rsid149 = new Rsid() { Val = "00FE2DFB" };
            Rsid rsid150 = new Rsid() { Val = "00FE4552" };
            Rsid rsid151 = new Rsid() { Val = "00FF6CCA" };

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
            W14.DocumentId documentId1 = new W14.DocumentId() { Val = "56B9EFC0" };
            W15.ChartTrackingRefBased chartTrackingRefBased1 = new W15.ChartTrackingRefBased();
            W15.PersistentDocumentId persistentDocumentId1 = new W15.PersistentDocumentId() { Val = "{058B6D5D-DDA3-4E00-B485-3115CB7DB2CC}" };

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

            Paragraph paragraph78 = new Paragraph() { RsidParagraphMarkRevision = "000E3880", RsidParagraphAddition = "000E3880", RsidParagraphProperties = "000E3880", RsidRunAdditionDefault = "000E3880", ParagraphId = "545481F4", TextId = "14B531FC" };

            ParagraphProperties paragraphProperties75 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId21 = new ParagraphStyleId() { Val = "a4" };

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Clear, Position = 8306 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 9752 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);

            paragraphProperties75.Append(paragraphStyleId21);
            paragraphProperties75.Append(tabs1);

            Run run176 = new Run();

            RunProperties runProperties176 = new RunProperties();
            RunFonts runFonts183 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color120 = new Color() { Val = "0000FF" };

            runProperties176.Append(runFonts183);
            runProperties176.Append(color120);
            Text text171 = new Text();
            text171.Text = dt.Rows[0]["CompanyName"].ToString();

            run176.Append(runProperties176);
            run176.Append(text171);

            Run run177 = new Run();

            RunProperties runProperties177 = new RunProperties();
            Color color121 = new Color() { Val = "0000FF" };

            runProperties177.Append(color121);
            Text text172 = new Text();
            text172.Text = "_";

            run177.Append(runProperties177);
            run177.Append(text172);

            Run run178 = new Run();

            RunProperties runProperties178 = new RunProperties();
            RunFonts runFonts184 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color122 = new Color() { Val = "0000FF" };

            runProperties178.Append(runFonts184);
            runProperties178.Append(color122);
            Text text173 = new Text();
            text173.Text = dt.Rows[0]["auditno"].ToString();

            run178.Append(runProperties178);
            run178.Append(text173);

            Run run179 = new Run();

            RunProperties runProperties179 = new RunProperties();
            Color color123 = new Color() { Val = "0000FF" };

            runProperties179.Append(color123);
            TabChar tabChar1 = new TabChar();

            run179.Append(runProperties179);
            run179.Append(tabChar1);

            Run run180 = new Run();

            RunProperties runProperties180 = new RunProperties();
            Color color124 = new Color() { Val = "0000FF" };

            runProperties180.Append(color124);
            TabChar tabChar2 = new TabChar();

            run180.Append(runProperties180);
            run180.Append(tabChar2);

            Run run181 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run181.Append(fieldChar1);

            Run run182 = new Run();
            FieldCode fieldCode1 = new FieldCode();
            fieldCode1.Text = "PAGE   \\* MERGEFORMAT";

            run182.Append(fieldCode1);

            Run run183 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run183.Append(fieldChar2);

            Run run184 = new Run() { RsidRunProperties = "00DB5532", RsidRunAddition = "00DB5532" };

            RunProperties runProperties181 = new RunProperties();
            NoProof noProof2 = new NoProof();
            Languages languages11 = new Languages() { Val = "zh-TW" };

            runProperties181.Append(noProof2);
            runProperties181.Append(languages11);
            Text text174 = new Text();
            text174.Text = "1";

            run184.Append(runProperties181);
            run184.Append(text174);

            Run run185 = new Run();
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run185.Append(fieldChar3);

            paragraph78.Append(paragraphProperties75);
            paragraph78.Append(run176);
            paragraph78.Append(run177);
            paragraph78.Append(run178);
            paragraph78.Append(run179);
            paragraph78.Append(run180);
            paragraph78.Append(run181);
            paragraph78.Append(run182);
            paragraph78.Append(run183);
            paragraph78.Append(run184);
            paragraph78.Append(run185);

            footer1.Append(paragraph78);

            footerPart1.Footer = footer1;
        }

        // Generates content of customXmlPart2.
        private void GenerateCustomXmlPart2Content(CustomXmlPart customXmlPart2)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart2.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?xml version=\"1.0\" encoding=\"utf-8\"?><p:properties xmlns:p=\"http://schemas.microsoft.com/office/2006/metadata/properties\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:pc=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"><documentManagement><lcf76f155ced4ddcb4097134ff3c332f xmlns=\"020e0566-15c6-45f6-927c-a14f1387091e\"><Terms xmlns=\"http://schemas.microsoft.com/office/infopath/2007/PartnerControls\"></Terms></lcf76f155ced4ddcb4097134ff3c332f><TaxCatchAll xmlns=\"cb5579a1-84df-45aa-8c56-deace17c375d\"/></documentManagement></p:properties>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart2.
        private void GenerateCustomXmlPropertiesPart2Content(CustomXmlPropertiesPart customXmlPropertiesPart2)
        {
            Ds.DataStoreItem dataStoreItem2 = new Ds.DataStoreItem() { ItemId = "{C488315B-259E-46F0-B46A-4026C844E630}" };
            dataStoreItem2.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences2 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference13 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/2006/metadata/properties" };
            Ds.SchemaReference schemaReference14 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/office/infopath/2007/PartnerControls" };
            Ds.SchemaReference schemaReference15 = new Ds.SchemaReference() { Uri = "020e0566-15c6-45f6-927c-a14f1387091e" };
            Ds.SchemaReference schemaReference16 = new Ds.SchemaReference() { Uri = "cb5579a1-84df-45aa-8c56-deace17c375d" };

            schemaReferences2.Append(schemaReference13);
            schemaReferences2.Append(schemaReference14);
            schemaReferences2.Append(schemaReference15);
            schemaReferences2.Append(schemaReference16);

            dataStoreItem2.Append(schemaReferences2);

            customXmlPropertiesPart2.DataStoreItem = dataStoreItem2;
        }

        // Generates content of customXmlPart3.
        private void GenerateCustomXmlPart3Content(CustomXmlPart customXmlPart3)
        {
            System.Xml.XmlTextWriter writer = new System.Xml.XmlTextWriter(customXmlPart3.GetStream(System.IO.FileMode.Create), System.Text.Encoding.UTF8);
            writer.WriteRaw("<?mso-contentType?><FormTemplates xmlns=\"http://schemas.microsoft.com/sharepoint/v3/contenttype/forms\"><Display>DocumentLibraryForm</Display><Edit>DocumentLibraryForm</Edit><New>DocumentLibraryForm</New></FormTemplates>");
            writer.Flush();
            writer.Close();
        }

        // Generates content of customXmlPropertiesPart3.
        private void GenerateCustomXmlPropertiesPart3Content(CustomXmlPropertiesPart customXmlPropertiesPart3)
        {
            Ds.DataStoreItem dataStoreItem3 = new Ds.DataStoreItem() { ItemId = "{FC031D83-AC92-4BC8-BDF2-74A420489E08}" };
            dataStoreItem3.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");

            Ds.SchemaReferences schemaReferences3 = new Ds.SchemaReferences();
            Ds.SchemaReference schemaReference17 = new Ds.SchemaReference() { Uri = "http://schemas.microsoft.com/sharepoint/v3/contenttype/forms" };

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
            RunFonts runFonts185 = new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "新細明體", ComplexScript = "Times New Roman" };
            Languages languages12 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts185);
            runPropertiesBaseStyle1.Append(languages12);

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
            Rsid rsid152 = new Rsid() { Val = "0014137D" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl() { Val = false };

            styleParagraphProperties1.Append(widowControl1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts186 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern2 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize218 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript71 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties1.Append(runFonts186);
            styleRunProperties1.Append(kern2);
            styleRunProperties1.Append(fontSize218);
            styleRunProperties1.Append(fontSizeComplexScript71);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid152);
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
            Rsid rsid153 = new Rsid() { Val = "0014137D" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs2.Append(tabStop3);
            tabs2.Append(tabStop4);
            SnapToGrid snapToGrid46 = new SnapToGrid() { Val = false };

            styleParagraphProperties2.Append(tabs2);
            styleParagraphProperties2.Append(snapToGrid46);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            FontSize fontSize219 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript72 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties2.Append(fontSize219);
            styleRunProperties2.Append(fontSizeComplexScript72);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(rsid153);
            style5.Append(styleParagraphProperties2);
            style5.Append(styleRunProperties2);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "a4" };
            StyleName styleName6 = new StyleName() { Val = "footer" };
            BasedOn basedOn2 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "a5" };
            UIPriority uIPriority4 = new UIPriority() { Val = 99 };
            Rsid rsid154 = new Rsid() { Val = "003C3877" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop5 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop6 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs3.Append(tabStop5);
            tabs3.Append(tabStop6);
            SnapToGrid snapToGrid47 = new SnapToGrid() { Val = false };

            styleParagraphProperties3.Append(tabs3);
            styleParagraphProperties3.Append(snapToGrid47);

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            FontSize fontSize220 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript73 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties3.Append(fontSize220);
            styleRunProperties3.Append(fontSizeComplexScript73);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle1);
            style6.Append(uIPriority4);
            style6.Append(rsid154);
            style6.Append(styleParagraphProperties3);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "a5", CustomStyle = true };
            StyleName styleName7 = new StyleName() { Val = "頁尾 字元" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a4" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            Rsid rsid155 = new Rsid() { Val = "003C3877" };

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            RunFonts runFonts187 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern3 = new Kern() { Val = (UInt32Value)2U };

            styleRunProperties4.Append(runFonts187);
            styleRunProperties4.Append(kern3);

            style7.Append(styleName7);
            style7.Append(linkedStyle2);
            style7.Append(uIPriority5);
            style7.Append(rsid155);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() { Type = StyleValues.Table, StyleId = "a6" };
            StyleName styleName8 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn3 = new BasedOn() { Val = "a1" };
            Rsid rsid156 = new Rsid() { Val = "00882386" };

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
            style8.Append(rsid156);
            style8.Append(styleTableProperties2);

            Style style9 = new Style() { Type = StyleValues.Paragraph, StyleId = "a7" };
            StyleName styleName9 = new StyleName() { Val = "footnote text" };
            BasedOn basedOn4 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "a8" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            Rsid rsid157 = new Rsid() { Val = "006C3779" };

            StyleParagraphProperties styleParagraphProperties4 = new StyleParagraphProperties();
            SnapToGrid snapToGrid48 = new SnapToGrid() { Val = false };

            styleParagraphProperties4.Append(snapToGrid48);

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            FontSize fontSize221 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript74 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties5.Append(fontSize221);
            styleRunProperties5.Append(fontSizeComplexScript74);

            style9.Append(styleName9);
            style9.Append(basedOn4);
            style9.Append(linkedStyle3);
            style9.Append(uIPriority6);
            style9.Append(rsid157);
            style9.Append(styleParagraphProperties4);
            style9.Append(styleRunProperties5);

            Style style10 = new Style() { Type = StyleValues.Character, StyleId = "a8", CustomStyle = true };
            StyleName styleName10 = new StyleName() { Val = "註腳文字 字元" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a7" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            Rsid rsid158 = new Rsid() { Val = "006C3779" };

            StyleRunProperties styleRunProperties6 = new StyleRunProperties();
            RunFonts runFonts188 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern4 = new Kern() { Val = (UInt32Value)2U };

            styleRunProperties6.Append(runFonts188);
            styleRunProperties6.Append(kern4);

            style10.Append(styleName10);
            style10.Append(linkedStyle4);
            style10.Append(uIPriority7);
            style10.Append(rsid158);
            style10.Append(styleRunProperties6);

            Style style11 = new Style() { Type = StyleValues.Character, StyleId = "a9" };
            StyleName styleName11 = new StyleName() { Val = "footnote reference" };
            Rsid rsid159 = new Rsid() { Val = "006C3779" };

            StyleRunProperties styleRunProperties7 = new StyleRunProperties();
            VerticalTextAlignment verticalTextAlignment1 = new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript };

            styleRunProperties7.Append(verticalTextAlignment1);

            style11.Append(styleName11);
            style11.Append(rsid159);
            style11.Append(styleRunProperties7);

            Style style12 = new Style() { Type = StyleValues.Character, StyleId = "aa" };
            StyleName styleName12 = new StyleName() { Val = "annotation reference" };
            Rsid rsid160 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties8 = new StyleRunProperties();
            FontSize fontSize222 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript75 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties8.Append(fontSize222);
            styleRunProperties8.Append(fontSizeComplexScript75);

            style12.Append(styleName12);
            style12.Append(rsid160);
            style12.Append(styleRunProperties8);

            Style style13 = new Style() { Type = StyleValues.Paragraph, StyleId = "ab" };
            StyleName styleName13 = new StyleName() { Val = "annotation text" };
            BasedOn basedOn5 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle5 = new LinkedStyle() { Val = "ac" };
            Rsid rsid161 = new Rsid() { Val = "00214C36" };

            style13.Append(styleName13);
            style13.Append(basedOn5);
            style13.Append(linkedStyle5);
            style13.Append(rsid161);

            Style style14 = new Style() { Type = StyleValues.Character, StyleId = "ac", CustomStyle = true };
            StyleName styleName14 = new StyleName() { Val = "註解文字 字元" };
            LinkedStyle linkedStyle6 = new LinkedStyle() { Val = "ab" };
            Rsid rsid162 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties9 = new StyleRunProperties();
            RunFonts runFonts189 = new RunFonts() { EastAsia = "標楷體" };
            Kern kern5 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize223 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript76 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties9.Append(runFonts189);
            styleRunProperties9.Append(kern5);
            styleRunProperties9.Append(fontSize223);
            styleRunProperties9.Append(fontSizeComplexScript76);

            style14.Append(styleName14);
            style14.Append(linkedStyle6);
            style14.Append(rsid162);
            style14.Append(styleRunProperties9);

            Style style15 = new Style() { Type = StyleValues.Paragraph, StyleId = "ad" };
            StyleName styleName15 = new StyleName() { Val = "annotation subject" };
            BasedOn basedOn6 = new BasedOn() { Val = "ab" };
            NextParagraphStyle nextParagraphStyle1 = new NextParagraphStyle() { Val = "ab" };
            LinkedStyle linkedStyle7 = new LinkedStyle() { Val = "ae" };
            Rsid rsid163 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties10 = new StyleRunProperties();
            Bold bold93 = new Bold();
            BoldComplexScript boldComplexScript55 = new BoldComplexScript();

            styleRunProperties10.Append(bold93);
            styleRunProperties10.Append(boldComplexScript55);

            style15.Append(styleName15);
            style15.Append(basedOn6);
            style15.Append(nextParagraphStyle1);
            style15.Append(linkedStyle7);
            style15.Append(rsid163);
            style15.Append(styleRunProperties10);

            Style style16 = new Style() { Type = StyleValues.Character, StyleId = "ae", CustomStyle = true };
            StyleName styleName16 = new StyleName() { Val = "註解主旨 字元" };
            LinkedStyle linkedStyle8 = new LinkedStyle() { Val = "ad" };
            Rsid rsid164 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties11 = new StyleRunProperties();
            RunFonts runFonts190 = new RunFonts() { EastAsia = "標楷體" };
            Bold bold94 = new Bold();
            BoldComplexScript boldComplexScript56 = new BoldComplexScript();
            Kern kern6 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize224 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript77 = new FontSizeComplexScript() { Val = "24" };

            styleRunProperties11.Append(runFonts190);
            styleRunProperties11.Append(bold94);
            styleRunProperties11.Append(boldComplexScript56);
            styleRunProperties11.Append(kern6);
            styleRunProperties11.Append(fontSize224);
            styleRunProperties11.Append(fontSizeComplexScript77);

            style16.Append(styleName16);
            style16.Append(linkedStyle8);
            style16.Append(rsid164);
            style16.Append(styleRunProperties11);

            Style style17 = new Style() { Type = StyleValues.Paragraph, StyleId = "af" };
            StyleName styleName17 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn7 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle9 = new LinkedStyle() { Val = "af0" };
            Rsid rsid165 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties12 = new StyleRunProperties();
            RunFonts runFonts191 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "新細明體" };
            FontSize fontSize225 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript78 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties12.Append(runFonts191);
            styleRunProperties12.Append(fontSize225);
            styleRunProperties12.Append(fontSizeComplexScript78);

            style17.Append(styleName17);
            style17.Append(basedOn7);
            style17.Append(linkedStyle9);
            style17.Append(rsid165);
            style17.Append(styleRunProperties12);

            Style style18 = new Style() { Type = StyleValues.Character, StyleId = "af0", CustomStyle = true };
            StyleName styleName18 = new StyleName() { Val = "註解方塊文字 字元" };
            LinkedStyle linkedStyle10 = new LinkedStyle() { Val = "af" };
            Rsid rsid166 = new Rsid() { Val = "00214C36" };

            StyleRunProperties styleRunProperties13 = new StyleRunProperties();
            RunFonts runFonts192 = new RunFonts() { Ascii = "Calibri Light", HighAnsi = "Calibri Light", EastAsia = "新細明體", ComplexScript = "Times New Roman" };
            Kern kern7 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize226 = new FontSize() { Val = "18" };
            FontSizeComplexScript fontSizeComplexScript79 = new FontSizeComplexScript() { Val = "18" };

            styleRunProperties13.Append(runFonts192);
            styleRunProperties13.Append(kern7);
            styleRunProperties13.Append(fontSize226);
            styleRunProperties13.Append(fontSizeComplexScript79);

            style18.Append(styleName18);
            style18.Append(linkedStyle10);
            style18.Append(rsid166);
            style18.Append(styleRunProperties13);

            Style style19 = new Style() { Type = StyleValues.Character, StyleId = "normaltextrun", CustomStyle = true };
            StyleName styleName19 = new StyleName() { Val = "normaltextrun" };
            Rsid rsid167 = new Rsid() { Val = "00EB5D9F" };

            style19.Append(styleName19);
            style19.Append(rsid167);

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

            Paragraph paragraph79 = new Paragraph() { RsidParagraphMarkRevision = "00213B89", RsidParagraphAddition = "00882386", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "008A0512", ParagraphId = "689B73A8", TextId = "5A7A805C" };

            ParagraphProperties paragraphProperties76 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId22 = new ParagraphStyleId() { Val = "a3" };

            paragraphProperties76.Append(paragraphStyleId22);

            Run run186 = new Run();

            RunProperties runProperties182 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties182.Append(noProof3);

            AlternateContent alternateContent2 = new AlternateContent();

            AlternateContentChoice alternateContentChoice2 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing2 = new Drawing();

            Wp.Anchor anchor2 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251657728U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true, EditId = "12B23A7D", AnchorId = "50319E5B" };
            Wp.SimplePosition simplePosition2 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition2 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset3 = new Wp.PositionOffset();
            positionOffset3.Text = "3086100";

            horizontalPosition2.Append(positionOffset3);

            Wp.VerticalPosition verticalPosition2 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset4 = new Wp.PositionOffset();
            positionOffset4.Text = "-100965";

            verticalPosition2.Append(positionOffset4);
            Wp.Extent extent2 = new Wp.Extent() { Cx = 3159760L, Cy = 356235L };
            Wp.EffectExtent effectExtent2 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone2 = new Wp.WrapNone();
            Wp.DocProperties docProperties2 = new Wp.DocProperties() { Id = (UInt32Value)9U, Name = "文字方塊 1" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties2 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks2 = new A.GraphicFrameLocks();
            graphicFrameLocks2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties2.Append(graphicFrameLocks2);

            A.Graphic graphic2 = new A.Graphic();
            graphic2.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData2 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape2 = new Wps.WordprocessingShape();

            Wps.NonVisualDrawingShapeProperties nonVisualDrawingShapeProperties2 = new Wps.NonVisualDrawingShapeProperties() { TextBox = true };
            A.ShapeLocks shapeLocks2 = new A.ShapeLocks();

            nonVisualDrawingShapeProperties2.Append(shapeLocks2);

            Wps.ShapeProperties shapeProperties2 = new Wps.ShapeProperties();

            A.Transform2D transform2D2 = new A.Transform2D();
            A.Offset offset2 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents2 = new A.Extents() { Cx = 3159760L, Cy = 356235L };

            transform2D2.Append(offset2);
            transform2D2.Append(extents2);

            A.PresetGeometry presetGeometry2 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList2 = new A.AdjustValueList();

            presetGeometry2.Append(adjustValueList2);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline2 = new A.Outline() { Width = 6350 };
            A.NoFill noFill2 = new A.NoFill();

            outline2.Append(noFill2);

            shapeProperties2.Append(transform2D2);
            shapeProperties2.Append(presetGeometry2);
            shapeProperties2.Append(noFill1);
            shapeProperties2.Append(outline2);

            Wps.TextBoxInfo2 textBoxInfo22 = new Wps.TextBoxInfo2();

            TextBoxContent textBoxContent3 = new TextBoxContent();

            Paragraph paragraph80 = new Paragraph() { RsidParagraphMarkRevision = "003A59C8", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "471EF4E2", TextId = "77777777" };

            ParagraphProperties paragraphProperties77 = new ParagraphProperties();
            SnapToGrid snapToGrid49 = new SnapToGrid() { Val = false };
            Justification justification28 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties70 = new ParagraphMarkRunProperties();
            FontSize fontSize227 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript80 = new FontSizeComplexScript() { Val = "20" };
            Languages languages13 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties70.Append(fontSize227);
            paragraphMarkRunProperties70.Append(fontSizeComplexScript80);
            paragraphMarkRunProperties70.Append(languages13);

            paragraphProperties77.Append(snapToGrid49);
            paragraphProperties77.Append(justification28);
            paragraphProperties77.Append(paragraphMarkRunProperties70);

            Run run187 = new Run() { RsidRunProperties = "003A59C8" };

            RunProperties runProperties183 = new RunProperties();
            FontSize fontSize228 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript81 = new FontSizeComplexScript() { Val = "20" };
            Languages languages14 = new Languages() { EastAsia = "zh-HK" };

            runProperties183.Append(fontSize228);
            runProperties183.Append(fontSizeComplexScript81);
            runProperties183.Append(languages14);
            Text text175 = new Text();
            text175.Text = "資安等級：機密";

            run187.Append(runProperties183);
            run187.Append(text175);

            paragraph80.Append(paragraphProperties77);
            paragraph80.Append(run187);

            Paragraph paragraph81 = new Paragraph() { RsidParagraphMarkRevision = "00F57CEA", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "09A559C5", TextId = "3BBBB6B9" };

            ParagraphProperties paragraphProperties78 = new ParagraphProperties();
            SnapToGrid snapToGrid50 = new SnapToGrid() { Val = false };
            Justification justification29 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties71 = new ParagraphMarkRunProperties();
            RunFonts runFonts193 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties71.Append(runFonts193);

            paragraphProperties78.Append(snapToGrid50);
            paragraphProperties78.Append(justification29);
            paragraphProperties78.Append(paragraphMarkRunProperties71);

            Run run188 = new Run() { RsidRunProperties = "7240FE02" };

            RunProperties runProperties184 = new RunProperties();
            FontSize fontSize229 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript82 = new FontSizeComplexScript() { Val = "20" };
            Languages languages15 = new Languages() { EastAsia = "zh-HK" };

            runProperties184.Append(fontSize229);
            runProperties184.Append(fontSizeComplexScript82);
            runProperties184.Append(languages15);
            Text text176 = new Text();
            text176.Text = "解密日期：";

            run188.Append(runProperties184);
            run188.Append(text176);

            Run run189 = new Run() { RsidRunProperties = "7240FE02" };

            RunProperties runProperties185 = new RunProperties();
            Color color125 = new Color() { Val = "3333FF" };
            FontSize fontSize230 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript83 = new FontSizeComplexScript() { Val = "20" };
            Languages languages16 = new Languages() { EastAsia = "zh-HK" };

            runProperties185.Append(color125);
            runProperties185.Append(fontSize230);
            runProperties185.Append(fontSizeComplexScript83);
            runProperties185.Append(languages16);
            Text text177 = new Text();
            if (DateTime.TryParse(dt.Rows[0]["enddate"].ToString(), out DateTime date4))
            {
                text177.Text = date4.ToString("yyyy-mm-dd");
            }
            else
            {
                text177.Text = "";
            }

            run189.Append(runProperties185);
            run189.Append(text177);

            Run run190 = new Run() { RsidRunAddition = "00F57CEA" };

            RunProperties runProperties186 = new RunProperties();
            RunFonts runFonts194 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color126 = new Color() { Val = "3333FF" };
            FontSize fontSize231 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript84 = new FontSizeComplexScript() { Val = "20" };
            Languages languages17 = new Languages() { EastAsia = "zh-HK" };

            runProperties186.Append(runFonts194);
            runProperties186.Append(color126);
            runProperties186.Append(fontSize231);
            runProperties186.Append(fontSizeComplexScript84);
            runProperties186.Append(languages17);
            Text text178 = new Text();
            text178.Text = "";

            run190.Append(runProperties186);
            run190.Append(text178);

            paragraph81.Append(paragraphProperties78);
            paragraph81.Append(run188);
            paragraph81.Append(run189);
            paragraph81.Append(run190);

            textBoxContent3.Append(paragraph80);
            textBoxContent3.Append(paragraph81);

            textBoxInfo22.Append(textBoxContent3);

            Wps.TextBodyProperties textBodyProperties2 = new Wps.TextBodyProperties() { Rotation = 0, UseParagraphSpacing = false, VerticalOverflow = A.TextVerticalOverflowValues.Overflow, HorizontalOverflow = A.TextHorizontalOverflowValues.Overflow, Vertical = A.TextVerticalValues.Horizontal, Wrap = A.TextWrappingValues.Square, LeftInset = 0, TopInset = 0, RightInset = 0, BottomInset = 0, ColumnCount = 1, ColumnSpacing = 0, RightToLeftColumns = false, FromWordArt = false, Anchor = A.TextAnchoringTypeValues.Top, AnchorCenter = false, ForceAntiAlias = false, CompatibleLineSpacing = true };

            A.PresetTextWrap presetTextWrap1 = new A.PresetTextWrap() { Preset = A.TextShapeValues.TextNoShape };
            A.AdjustValueList adjustValueList3 = new A.AdjustValueList();

            presetTextWrap1.Append(adjustValueList3);
            A.NoAutoFit noAutoFit2 = new A.NoAutoFit();

            textBodyProperties2.Append(presetTextWrap1);
            textBodyProperties2.Append(noAutoFit2);

            wordprocessingShape2.Append(nonVisualDrawingShapeProperties2);
            wordprocessingShape2.Append(shapeProperties2);
            wordprocessingShape2.Append(textBoxInfo22);
            wordprocessingShape2.Append(textBodyProperties2);

            graphicData2.Append(wordprocessingShape2);

            graphic2.Append(graphicData2);

            Wp14.RelativeWidth relativeWidth2 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin };
            Wp14.PercentageWidth percentageWidth2 = new Wp14.PercentageWidth();
            percentageWidth2.Text = "0";

            relativeWidth2.Append(percentageWidth2);

            Wp14.RelativeHeight relativeHeight2 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin };
            Wp14.PercentageHeight percentageHeight2 = new Wp14.PercentageHeight();
            percentageHeight2.Text = "0";

            relativeHeight2.Append(percentageHeight2);

            anchor2.Append(simplePosition2);
            anchor2.Append(horizontalPosition2);
            anchor2.Append(verticalPosition2);
            anchor2.Append(extent2);
            anchor2.Append(effectExtent2);
            anchor2.Append(wrapNone2);
            anchor2.Append(docProperties2);
            anchor2.Append(nonVisualGraphicFrameDrawingProperties2);
            anchor2.Append(graphic2);
            anchor2.Append(relativeWidth2);
            anchor2.Append(relativeHeight2);

            drawing2.Append(anchor2);

            alternateContentChoice2.Append(drawing2);

            AlternateContentFallback alternateContentFallback2 = new AlternateContentFallback();

            Picture picture2 = new Picture();

            V.Shapetype shapetype2 = new V.Shapetype() { Id = "_x0000_t202", CoordinateSize = "21600,21600", OptionalNumber = 202, EdgePath = "m,l,21600r21600,l21600,xe" };
            shapetype2.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "50319E5B"));
            V.Stroke stroke3 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            V.Path path2 = new V.Path() { AllowGradientShape = true, ConnectionPointType = Ovml.ConnectValues.Rectangle };

            shapetype2.Append(stroke3);
            shapetype2.Append(path2);

            V.Shape shape2 = new V.Shape() { Id = "文字方塊 1", Style = "position:absolute;margin-left:243pt;margin-top:-7.95pt;width:248.8pt;height:28.05pt;z-index:251657728;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top", OptionalString = "_x0000_s1027", Filled = false, Stroked = false, StrokeWeight = ".5pt", Type = "#_x0000_t202", EncodedPackage = "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQAzVqvWGAIAADUEAAAOAAAAZHJzL2Uyb0RvYy54bWysU11v0zAUfUfiP1h+p+mHWljUdCqbipCq\nbVKH9uw6dhvh+Jprt8n49Vw7SYsGT4gX58b3+5zj5W1bG3ZW6CuwBZ+MxpwpK6Gs7KHg3543Hz5x\n5oOwpTBgVcFflee3q/fvlo3L1RSOYEqFjIpYnzeu4McQXJ5lXh5VLfwInLLk1IC1CPSLh6xE0VD1\n2mTT8XiRNYClQ5DKe7q975x8leprrWR41NqrwEzBabaQTkznPp7ZainyAwp3rGQ/hviHKWpRWWp6\nKXUvgmAnrP4oVVcSwYMOIwl1BlpXUqUdaJvJ+M02u6NwKu1C4Hh3gcn/v7Ly4bxzT8hC+xlaIjAt\n4d0W5HdP2GSN83kfEzH1uafouGirsY5fWoFRImH7esFTtYFJupxN5jcfF+SS5JvNF9PZPAKeXbMd\n+vBFQc2iUXAkvtIE4rz1oQsdQmIzC5vKmMSZsawp+GI2H6eEi4eKG9sP3s0apw7tvqW0aO6hfKWF\nEToteCc3FTXfCh+eBBL5NC8JOjzSoQ1QE+gtzo6AP/92H+OJE/Jy1pCYCu5/nAQqzsxXS2xF5Q0G\nDsZ+MOypvgPS54SeipPJpAQMZjA1Qv1COl/HLuQSVlKvgofBvAudpOmdSLVepyDSlxNha3dODrxG\nKJ/bF4GuxzsQUw8wyEzkb2DvYjvg16cAukqcXFHscSZtJlb7dxTF//t/irq+9tUvAAAA//8DAFBL\nAwQUAAYACAAAACEAgOgf6eAAAAAKAQAADwAAAGRycy9kb3ducmV2LnhtbEyPzU7DMBCE70i8g7VI\n3Fo7BaI0xKkQPzcotAUJbk5skojYjuxNGt6e5QTH0Yxmvik2s+3ZZELsvJOQLAUw42qvO9dIeD08\nLDJgEZXTqvfOSPg2ETbl6Umhcu2PbmemPTaMSlzMlYQWccg5j3VrrIpLPxhH3qcPViHJ0HAd1JHK\nbc9XQqTcqs7RQqsGc9ua+ms/Wgn9ewyPlcCP6a55wpdnPr7dJ1spz8/mm2tgaGb8C8MvPqFDSUyV\nH52OrJdwmaX0BSUskqs1MEqss4sUWEWWWAEvC/7/QvkDAAD//wMAUEsBAi0AFAAGAAgAAAAhALaD\nOJL+AAAA4QEAABMAAAAAAAAAAAAAAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYA\nCAAAACEAOP0h/9YAAACUAQAACwAAAAAAAAAAAAAAAAAvAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYA\nCAAAACEAM1ar1hgCAAA1BAAADgAAAAAAAAAAAAAAAAAuAgAAZHJzL2Uyb0RvYy54bWxQSwECLQAU\nAAYACAAAACEAgOgf6eAAAAAKAQAADwAAAAAAAAAAAAAAAAByBAAAZHJzL2Rvd25yZXYueG1sUEsF\nBgAAAAAEAAQA8wAAAH8FAAAAAA==\n" };

            V.TextBox textBox2 = new V.TextBox() { Inset = "0,0,0,0" };

            TextBoxContent textBoxContent4 = new TextBoxContent();

            Paragraph paragraph82 = new Paragraph() { RsidParagraphMarkRevision = "003A59C8", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "471EF4E2", TextId = "77777777" };

            ParagraphProperties paragraphProperties79 = new ParagraphProperties();
            SnapToGrid snapToGrid51 = new SnapToGrid() { Val = false };
            Justification justification30 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties72 = new ParagraphMarkRunProperties();
            FontSize fontSize232 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript85 = new FontSizeComplexScript() { Val = "20" };
            Languages languages18 = new Languages() { EastAsia = "zh-HK" };

            paragraphMarkRunProperties72.Append(fontSize232);
            paragraphMarkRunProperties72.Append(fontSizeComplexScript85);
            paragraphMarkRunProperties72.Append(languages18);

            paragraphProperties79.Append(snapToGrid51);
            paragraphProperties79.Append(justification30);
            paragraphProperties79.Append(paragraphMarkRunProperties72);

            Run run191 = new Run() { RsidRunProperties = "003A59C8" };

            RunProperties runProperties187 = new RunProperties();
            FontSize fontSize233 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript86 = new FontSizeComplexScript() { Val = "20" };
            Languages languages19 = new Languages() { EastAsia = "zh-HK" };

            runProperties187.Append(fontSize233);
            runProperties187.Append(fontSizeComplexScript86);
            runProperties187.Append(languages19);
            Text text179 = new Text();
            text179.Text = "資安等級：機密";

            run191.Append(runProperties187);
            run191.Append(text179);

            paragraph82.Append(paragraphProperties79);
            paragraph82.Append(run191);

            Paragraph paragraph83 = new Paragraph() { RsidParagraphMarkRevision = "00F57CEA", RsidParagraphAddition = "00213B89", RsidParagraphProperties = "00213B89", RsidRunAdditionDefault = "00213B89", ParagraphId = "09A559C5", TextId = "3BBBB6B9" };

            ParagraphProperties paragraphProperties80 = new ParagraphProperties();
            SnapToGrid snapToGrid52 = new SnapToGrid() { Val = false };
            Justification justification31 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties73 = new ParagraphMarkRunProperties();
            RunFonts runFonts195 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties73.Append(runFonts195);

            paragraphProperties80.Append(snapToGrid52);
            paragraphProperties80.Append(justification31);
            paragraphProperties80.Append(paragraphMarkRunProperties73);

            Run run192 = new Run() { RsidRunProperties = "7240FE02" };

            RunProperties runProperties188 = new RunProperties();
            FontSize fontSize234 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript87 = new FontSizeComplexScript() { Val = "20" };
            Languages languages20 = new Languages() { EastAsia = "zh-HK" };

            runProperties188.Append(fontSize234);
            runProperties188.Append(fontSizeComplexScript87);
            runProperties188.Append(languages20);
            Text text180 = new Text();
            text180.Text = "解密日期：";

            run192.Append(runProperties188);
            run192.Append(text180);

            Run run193 = new Run() { RsidRunProperties = "7240FE02" };

            RunProperties runProperties189 = new RunProperties();
            Color color127 = new Color() { Val = "3333FF" };
            FontSize fontSize235 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript88 = new FontSizeComplexScript() { Val = "20" };
            Languages languages21 = new Languages() { EastAsia = "zh-HK" };

            runProperties189.Append(color127);
            runProperties189.Append(fontSize235);
            runProperties189.Append(fontSizeComplexScript88);
            runProperties189.Append(languages21);
            Text text181 = new Text();
    
            if (DateTime.TryParse(dt.Rows[0]["Ptitleenddate"].ToString(), out DateTime date5))
            {
                text181.Text = date5.ToString("yyyy-mm-dd");
            }
            else
            {
                text181.Text = "";
            }

            run193.Append(runProperties189);
            run193.Append(text181);

            Run run194 = new Run() { RsidRunAddition = "00F57CEA" };

            RunProperties runProperties190 = new RunProperties();
            RunFonts runFonts196 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Color color128 = new Color() { Val = "3333FF" };
            FontSize fontSize236 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript89 = new FontSizeComplexScript() { Val = "20" };
            Languages languages22 = new Languages() { EastAsia = "zh-HK" };

            runProperties190.Append(runFonts196);
            runProperties190.Append(color128);
            runProperties190.Append(fontSize236);
            runProperties190.Append(fontSizeComplexScript89);
            runProperties190.Append(languages22);
            Text text182 = new Text();
            text182.Text = "";

            run194.Append(runProperties190);
            run194.Append(text182);

            paragraph83.Append(paragraphProperties80);
            paragraph83.Append(run192);
            paragraph83.Append(run193);
            paragraph83.Append(run194);

            textBoxContent4.Append(paragraph82);
            textBoxContent4.Append(paragraph83);

            textBox2.Append(textBoxContent4);

            shape2.Append(textBox2);

            picture2.Append(shapetype2);
            picture2.Append(shape2);

            alternateContentFallback2.Append(picture2);

            alternateContent2.Append(alternateContentChoice2);
            alternateContent2.Append(alternateContentFallback2);

            run186.Append(runProperties182);
            run186.Append(alternateContent2);

            paragraph79.Append(paragraphProperties76);
            paragraph79.Append(run186);

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
            Nsid nsid1 = new Nsid() { Val = "0D104B72" };
            MultiLevelType multiLevelType1 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode1 = new TemplateCode() { Val = "24FA0A5A" };

            Level level1 = new Level() { LevelIndex = 0, TemplateCode = "04090015" };
            StartNumberingValue startNumberingValue1 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat1 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText1 = new LevelText() { Val = "%1、" };
            LevelJustification levelJustification1 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties1 = new PreviousParagraphProperties();
            Indentation indentation39 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties1.Append(indentation39);

            level1.Append(startNumberingValue1);
            level1.Append(numberingFormat1);
            level1.Append(levelText1);
            level1.Append(levelJustification1);
            level1.Append(previousParagraphProperties1);

            Level level2 = new Level() { LevelIndex = 1, TemplateCode = "164A8326" };
            StartNumberingValue startNumberingValue2 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat2 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText2 = new LevelText() { Val = "(%2)" };
            LevelJustification levelJustification2 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties2 = new PreviousParagraphProperties();
            Indentation indentation40 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties2.Append(indentation40);

            NumberingSymbolRunProperties numberingSymbolRunProperties1 = new NumberingSymbolRunProperties();
            RunFonts runFonts197 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Bold bold95 = new Bold() { Val = false };
            Italic italic1 = new Italic() { Val = false };

            numberingSymbolRunProperties1.Append(runFonts197);
            numberingSymbolRunProperties1.Append(bold95);
            numberingSymbolRunProperties1.Append(italic1);

            level2.Append(startNumberingValue2);
            level2.Append(numberingFormat2);
            level2.Append(levelText2);
            level2.Append(levelJustification2);
            level2.Append(previousParagraphProperties2);
            level2.Append(numberingSymbolRunProperties1);

            Level level3 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue3 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat3 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText3 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification3 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties3 = new PreviousParagraphProperties();
            Indentation indentation41 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties3.Append(indentation41);

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
            Indentation indentation42 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties4.Append(indentation42);

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
            Indentation indentation43 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties5.Append(indentation43);

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
            Indentation indentation44 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties6.Append(indentation44);

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
            Indentation indentation45 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties7.Append(indentation45);

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
            Indentation indentation46 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties8.Append(indentation46);

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
            Indentation indentation47 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties9.Append(indentation47);

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
            Nsid nsid2 = new Nsid() { Val = "0D330787" };
            MultiLevelType multiLevelType2 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode2 = new TemplateCode() { Val = "F13E9CF0" };

            Level level10 = new Level() { LevelIndex = 0, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue10 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat10 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText10 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification10 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties10 = new PreviousParagraphProperties();
            Indentation indentation48 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties10.Append(indentation48);

            NumberingSymbolRunProperties numberingSymbolRunProperties2 = new NumberingSymbolRunProperties();
            RunFonts runFonts198 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties2.Append(runFonts198);

            level10.Append(startNumberingValue10);
            level10.Append(numberingFormat10);
            level10.Append(levelText10);
            level10.Append(levelJustification10);
            level10.Append(previousParagraphProperties10);
            level10.Append(numberingSymbolRunProperties2);

            Level level11 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue11 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat11 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText11 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification11 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties11 = new PreviousParagraphProperties();
            Indentation indentation49 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties11.Append(indentation49);

            level11.Append(startNumberingValue11);
            level11.Append(numberingFormat11);
            level11.Append(levelText11);
            level11.Append(levelJustification11);
            level11.Append(previousParagraphProperties11);

            Level level12 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue12 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat12 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText12 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification12 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties12 = new PreviousParagraphProperties();
            Indentation indentation50 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties12.Append(indentation50);

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
            Indentation indentation51 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties13.Append(indentation51);

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
            Indentation indentation52 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties14.Append(indentation52);

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
            Indentation indentation53 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties15.Append(indentation53);

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
            Indentation indentation54 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties16.Append(indentation54);

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
            Indentation indentation55 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties17.Append(indentation55);

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
            Indentation indentation56 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties18.Append(indentation56);

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
            Nsid nsid3 = new Nsid() { Val = "1B914155" };
            MultiLevelType multiLevelType3 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode3 = new TemplateCode() { Val = "2286B7EA" };

            Level level19 = new Level() { LevelIndex = 0, TemplateCode = "04090001" };
            StartNumberingValue startNumberingValue19 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat19 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText19 = new LevelText() { Val = "l" };
            LevelJustification levelJustification19 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties19 = new PreviousParagraphProperties();
            Indentation indentation57 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties19.Append(indentation57);

            NumberingSymbolRunProperties numberingSymbolRunProperties3 = new NumberingSymbolRunProperties();
            RunFonts runFonts199 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties3.Append(runFonts199);

            level19.Append(startNumberingValue19);
            level19.Append(numberingFormat19);
            level19.Append(levelText19);
            level19.Append(levelJustification19);
            level19.Append(previousParagraphProperties19);
            level19.Append(numberingSymbolRunProperties3);

            Level level20 = new Level() { LevelIndex = 1, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue20 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat20 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText20 = new LevelText() { Val = "n" };
            LevelJustification levelJustification20 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties20 = new PreviousParagraphProperties();
            Indentation indentation58 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties20.Append(indentation58);

            NumberingSymbolRunProperties numberingSymbolRunProperties4 = new NumberingSymbolRunProperties();
            RunFonts runFonts200 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties4.Append(runFonts200);

            level20.Append(startNumberingValue20);
            level20.Append(numberingFormat20);
            level20.Append(levelText20);
            level20.Append(levelJustification20);
            level20.Append(previousParagraphProperties20);
            level20.Append(numberingSymbolRunProperties4);

            Level level21 = new Level() { LevelIndex = 2, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue21 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat21 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText21 = new LevelText() { Val = "u" };
            LevelJustification levelJustification21 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties21 = new PreviousParagraphProperties();
            Indentation indentation59 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties21.Append(indentation59);

            NumberingSymbolRunProperties numberingSymbolRunProperties5 = new NumberingSymbolRunProperties();
            RunFonts runFonts201 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties5.Append(runFonts201);

            level21.Append(startNumberingValue21);
            level21.Append(numberingFormat21);
            level21.Append(levelText21);
            level21.Append(levelJustification21);
            level21.Append(previousParagraphProperties21);
            level21.Append(numberingSymbolRunProperties5);

            Level level22 = new Level() { LevelIndex = 3, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue22 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat22 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText22 = new LevelText() { Val = "l" };
            LevelJustification levelJustification22 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties22 = new PreviousParagraphProperties();
            Indentation indentation60 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties22.Append(indentation60);

            NumberingSymbolRunProperties numberingSymbolRunProperties6 = new NumberingSymbolRunProperties();
            RunFonts runFonts202 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties6.Append(runFonts202);

            level22.Append(startNumberingValue22);
            level22.Append(numberingFormat22);
            level22.Append(levelText22);
            level22.Append(levelJustification22);
            level22.Append(previousParagraphProperties22);
            level22.Append(numberingSymbolRunProperties6);

            Level level23 = new Level() { LevelIndex = 4, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue23 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat23 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText23 = new LevelText() { Val = "n" };
            LevelJustification levelJustification23 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties23 = new PreviousParagraphProperties();
            Indentation indentation61 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties23.Append(indentation61);

            NumberingSymbolRunProperties numberingSymbolRunProperties7 = new NumberingSymbolRunProperties();
            RunFonts runFonts203 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties7.Append(runFonts203);

            level23.Append(startNumberingValue23);
            level23.Append(numberingFormat23);
            level23.Append(levelText23);
            level23.Append(levelJustification23);
            level23.Append(previousParagraphProperties23);
            level23.Append(numberingSymbolRunProperties7);

            Level level24 = new Level() { LevelIndex = 5, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue24 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat24 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText24 = new LevelText() { Val = "u" };
            LevelJustification levelJustification24 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties24 = new PreviousParagraphProperties();
            Indentation indentation62 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties24.Append(indentation62);

            NumberingSymbolRunProperties numberingSymbolRunProperties8 = new NumberingSymbolRunProperties();
            RunFonts runFonts204 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties8.Append(runFonts204);

            level24.Append(startNumberingValue24);
            level24.Append(numberingFormat24);
            level24.Append(levelText24);
            level24.Append(levelJustification24);
            level24.Append(previousParagraphProperties24);
            level24.Append(numberingSymbolRunProperties8);

            Level level25 = new Level() { LevelIndex = 6, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue25 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat25 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText25 = new LevelText() { Val = "l" };
            LevelJustification levelJustification25 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties25 = new PreviousParagraphProperties();
            Indentation indentation63 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties25.Append(indentation63);

            NumberingSymbolRunProperties numberingSymbolRunProperties9 = new NumberingSymbolRunProperties();
            RunFonts runFonts205 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties9.Append(runFonts205);

            level25.Append(startNumberingValue25);
            level25.Append(numberingFormat25);
            level25.Append(levelText25);
            level25.Append(levelJustification25);
            level25.Append(previousParagraphProperties25);
            level25.Append(numberingSymbolRunProperties9);

            Level level26 = new Level() { LevelIndex = 7, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue26 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat26 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText26 = new LevelText() { Val = "n" };
            LevelJustification levelJustification26 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties26 = new PreviousParagraphProperties();
            Indentation indentation64 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties26.Append(indentation64);

            NumberingSymbolRunProperties numberingSymbolRunProperties10 = new NumberingSymbolRunProperties();
            RunFonts runFonts206 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties10.Append(runFonts206);

            level26.Append(startNumberingValue26);
            level26.Append(numberingFormat26);
            level26.Append(levelText26);
            level26.Append(levelJustification26);
            level26.Append(previousParagraphProperties26);
            level26.Append(numberingSymbolRunProperties10);

            Level level27 = new Level() { LevelIndex = 8, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue27 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat27 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText27 = new LevelText() { Val = "u" };
            LevelJustification levelJustification27 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties27 = new PreviousParagraphProperties();
            Indentation indentation65 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties27.Append(indentation65);

            NumberingSymbolRunProperties numberingSymbolRunProperties11 = new NumberingSymbolRunProperties();
            RunFonts runFonts207 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties11.Append(runFonts207);

            level27.Append(startNumberingValue27);
            level27.Append(numberingFormat27);
            level27.Append(levelText27);
            level27.Append(levelJustification27);
            level27.Append(previousParagraphProperties27);
            level27.Append(numberingSymbolRunProperties11);

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
            Nsid nsid4 = new Nsid() { Val = "20EE0C91" };
            MultiLevelType multiLevelType4 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode4 = new TemplateCode() { Val = "6FA82208" };

            Level level28 = new Level() { LevelIndex = 0, TemplateCode = "04090017" };
            StartNumberingValue startNumberingValue28 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat28 = new NumberingFormat() { Val = NumberFormatValues.IdeographLegalTraditional };
            LevelText levelText28 = new LevelText() { Val = "%1、" };
            LevelJustification levelJustification28 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties28 = new PreviousParagraphProperties();

            Tabs tabs4 = new Tabs();
            TabStop tabStop7 = new TabStop() { Val = TabStopValues.Number, Position = 482 };

            tabs4.Append(tabStop7);
            Indentation indentation66 = new Indentation() { Start = "482", Hanging = "480" };

            previousParagraphProperties28.Append(tabs4);
            previousParagraphProperties28.Append(indentation66);

            NumberingSymbolRunProperties numberingSymbolRunProperties12 = new NumberingSymbolRunProperties();
            RunFonts runFonts208 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties12.Append(runFonts208);

            level28.Append(startNumberingValue28);
            level28.Append(numberingFormat28);
            level28.Append(levelText28);
            level28.Append(levelJustification28);
            level28.Append(previousParagraphProperties28);
            level28.Append(numberingSymbolRunProperties12);

            Level level29 = new Level() { LevelIndex = 1, TemplateCode = "738AFC8E" };
            StartNumberingValue startNumberingValue29 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat29 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText29 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification29 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties29 = new PreviousParagraphProperties();

            Tabs tabs5 = new Tabs();
            TabStop tabStop8 = new TabStop() { Val = TabStopValues.Number, Position = 1082 };

            tabs5.Append(tabStop8);
            Indentation indentation67 = new Indentation() { Start = "1082", Hanging = "600" };

            previousParagraphProperties29.Append(tabs5);
            previousParagraphProperties29.Append(indentation67);

            NumberingSymbolRunProperties numberingSymbolRunProperties13 = new NumberingSymbolRunProperties();
            RunFonts runFonts209 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            Spacing spacing1 = new Spacing() { Val = -10 };
            Position position1 = new Position() { Val = "0" };

            numberingSymbolRunProperties13.Append(runFonts209);
            numberingSymbolRunProperties13.Append(spacing1);
            numberingSymbolRunProperties13.Append(position1);

            level29.Append(startNumberingValue29);
            level29.Append(numberingFormat29);
            level29.Append(levelText29);
            level29.Append(levelJustification29);
            level29.Append(previousParagraphProperties29);
            level29.Append(numberingSymbolRunProperties13);

            Level level30 = new Level() { LevelIndex = 2, TemplateCode = "7AD827AC" };
            StartNumberingValue startNumberingValue30 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat30 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText30 = new LevelText() { Val = "(%3)" };
            LevelJustification levelJustification30 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties30 = new PreviousParagraphProperties();

            Tabs tabs6 = new Tabs();
            TabStop tabStop9 = new TabStop() { Val = TabStopValues.Number, Position = 1701 };

            tabs6.Append(tabStop9);
            Indentation indentation68 = new Indentation() { Start = "1701", Hanging = "624" };

            previousParagraphProperties30.Append(tabs6);
            previousParagraphProperties30.Append(indentation68);

            NumberingSymbolRunProperties numberingSymbolRunProperties14 = new NumberingSymbolRunProperties();
            RunFonts runFonts210 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, EastAsia = "標楷體" };
            Bold bold96 = new Bold() { Val = false };
            Italic italic2 = new Italic() { Val = false };
            Spacing spacing2 = new Spacing() { Val = -10 };
            Position position2 = new Position() { Val = "0" };

            numberingSymbolRunProperties14.Append(runFonts210);
            numberingSymbolRunProperties14.Append(bold96);
            numberingSymbolRunProperties14.Append(italic2);
            numberingSymbolRunProperties14.Append(spacing2);
            numberingSymbolRunProperties14.Append(position2);

            level30.Append(startNumberingValue30);
            level30.Append(numberingFormat30);
            level30.Append(levelText30);
            level30.Append(levelJustification30);
            level30.Append(previousParagraphProperties30);
            level30.Append(numberingSymbolRunProperties14);

            Level level31 = new Level() { LevelIndex = 3, TemplateCode = "6814211C" };
            StartNumberingValue startNumberingValue31 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat31 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText31 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification31 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties31 = new PreviousParagraphProperties();

            Tabs tabs7 = new Tabs();
            TabStop tabStop10 = new TabStop() { Val = TabStopValues.Number, Position = 2041 };

            tabs7.Append(tabStop10);
            Indentation indentation69 = new Indentation() { Start = "2041", Hanging = "340" };

            previousParagraphProperties31.Append(tabs7);
            previousParagraphProperties31.Append(indentation69);

            NumberingSymbolRunProperties numberingSymbolRunProperties15 = new NumberingSymbolRunProperties();
            RunFonts runFonts211 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties15.Append(runFonts211);

            level31.Append(startNumberingValue31);
            level31.Append(numberingFormat31);
            level31.Append(levelText31);
            level31.Append(levelJustification31);
            level31.Append(previousParagraphProperties31);
            level31.Append(numberingSymbolRunProperties15);

            Level level32 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue32 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat32 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText32 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification32 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties32 = new PreviousParagraphProperties();

            Tabs tabs8 = new Tabs();
            TabStop tabStop11 = new TabStop() { Val = TabStopValues.Number, Position = 2402 };

            tabs8.Append(tabStop11);
            Indentation indentation70 = new Indentation() { Start = "2402", Hanging = "480" };

            previousParagraphProperties32.Append(tabs8);
            previousParagraphProperties32.Append(indentation70);

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

            Tabs tabs9 = new Tabs();
            TabStop tabStop12 = new TabStop() { Val = TabStopValues.Number, Position = 2882 };

            tabs9.Append(tabStop12);
            Indentation indentation71 = new Indentation() { Start = "2882", Hanging = "480" };

            previousParagraphProperties33.Append(tabs9);
            previousParagraphProperties33.Append(indentation71);

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

            Tabs tabs10 = new Tabs();
            TabStop tabStop13 = new TabStop() { Val = TabStopValues.Number, Position = 3362 };

            tabs10.Append(tabStop13);
            Indentation indentation72 = new Indentation() { Start = "3362", Hanging = "480" };

            previousParagraphProperties34.Append(tabs10);
            previousParagraphProperties34.Append(indentation72);

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

            Tabs tabs11 = new Tabs();
            TabStop tabStop14 = new TabStop() { Val = TabStopValues.Number, Position = 3842 };

            tabs11.Append(tabStop14);
            Indentation indentation73 = new Indentation() { Start = "3842", Hanging = "480" };

            previousParagraphProperties35.Append(tabs11);
            previousParagraphProperties35.Append(indentation73);

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

            Tabs tabs12 = new Tabs();
            TabStop tabStop15 = new TabStop() { Val = TabStopValues.Number, Position = 4322 };

            tabs12.Append(tabStop15);
            Indentation indentation74 = new Indentation() { Start = "4322", Hanging = "480" };

            previousParagraphProperties36.Append(tabs12);
            previousParagraphProperties36.Append(indentation74);

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
            Nsid nsid5 = new Nsid() { Val = "2DF2620B" };
            MultiLevelType multiLevelType5 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode5 = new TemplateCode() { Val = "93440D90" };

            Level level37 = new Level() { LevelIndex = 0, TemplateCode = "AA78403C" };
            StartNumberingValue startNumberingValue37 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat37 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText37 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification37 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties37 = new PreviousParagraphProperties();
            Indentation indentation75 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties37.Append(indentation75);

            NumberingSymbolRunProperties numberingSymbolRunProperties16 = new NumberingSymbolRunProperties();
            RunFonts runFonts212 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties16.Append(runFonts212);

            level37.Append(startNumberingValue37);
            level37.Append(numberingFormat37);
            level37.Append(levelText37);
            level37.Append(levelJustification37);
            level37.Append(previousParagraphProperties37);
            level37.Append(numberingSymbolRunProperties16);

            Level level38 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue38 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat38 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText38 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification38 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties38 = new PreviousParagraphProperties();
            Indentation indentation76 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties38.Append(indentation76);

            level38.Append(startNumberingValue38);
            level38.Append(numberingFormat38);
            level38.Append(levelText38);
            level38.Append(levelJustification38);
            level38.Append(previousParagraphProperties38);

            Level level39 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue39 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat39 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText39 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification39 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties39 = new PreviousParagraphProperties();
            Indentation indentation77 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties39.Append(indentation77);

            level39.Append(startNumberingValue39);
            level39.Append(numberingFormat39);
            level39.Append(levelText39);
            level39.Append(levelJustification39);
            level39.Append(previousParagraphProperties39);

            Level level40 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue40 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat40 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText40 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification40 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties40 = new PreviousParagraphProperties();
            Indentation indentation78 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties40.Append(indentation78);

            level40.Append(startNumberingValue40);
            level40.Append(numberingFormat40);
            level40.Append(levelText40);
            level40.Append(levelJustification40);
            level40.Append(previousParagraphProperties40);

            Level level41 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue41 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat41 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText41 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification41 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties41 = new PreviousParagraphProperties();
            Indentation indentation79 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties41.Append(indentation79);

            level41.Append(startNumberingValue41);
            level41.Append(numberingFormat41);
            level41.Append(levelText41);
            level41.Append(levelJustification41);
            level41.Append(previousParagraphProperties41);

            Level level42 = new Level() { LevelIndex = 5, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue42 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat42 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText42 = new LevelText() { Val = "%6." };
            LevelJustification levelJustification42 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties42 = new PreviousParagraphProperties();
            Indentation indentation80 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties42.Append(indentation80);

            level42.Append(startNumberingValue42);
            level42.Append(numberingFormat42);
            level42.Append(levelText42);
            level42.Append(levelJustification42);
            level42.Append(previousParagraphProperties42);

            Level level43 = new Level() { LevelIndex = 6, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue43 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat43 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText43 = new LevelText() { Val = "%7." };
            LevelJustification levelJustification43 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties43 = new PreviousParagraphProperties();
            Indentation indentation81 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties43.Append(indentation81);

            level43.Append(startNumberingValue43);
            level43.Append(numberingFormat43);
            level43.Append(levelText43);
            level43.Append(levelJustification43);
            level43.Append(previousParagraphProperties43);

            Level level44 = new Level() { LevelIndex = 7, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue44 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat44 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText44 = new LevelText() { Val = "%8、" };
            LevelJustification levelJustification44 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties44 = new PreviousParagraphProperties();
            Indentation indentation82 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties44.Append(indentation82);

            level44.Append(startNumberingValue44);
            level44.Append(numberingFormat44);
            level44.Append(levelText44);
            level44.Append(levelJustification44);
            level44.Append(previousParagraphProperties44);

            Level level45 = new Level() { LevelIndex = 8, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue45 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat45 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText45 = new LevelText() { Val = "%9." };
            LevelJustification levelJustification45 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties45 = new PreviousParagraphProperties();
            Indentation indentation83 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties45.Append(indentation83);

            level45.Append(startNumberingValue45);
            level45.Append(numberingFormat45);
            level45.Append(levelText45);
            level45.Append(levelJustification45);
            level45.Append(previousParagraphProperties45);

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
            Nsid nsid6 = new Nsid() { Val = "3AD5471D" };
            MultiLevelType multiLevelType6 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode6 = new TemplateCode() { Val = "D58AC6F4" };

            Level level46 = new Level() { LevelIndex = 0, TemplateCode = "1046A65E" };
            StartNumberingValue startNumberingValue46 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat46 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText46 = new LevelText() { Val = "(%1)" };
            LevelJustification levelJustification46 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties46 = new PreviousParagraphProperties();

            Tabs tabs13 = new Tabs();
            TabStop tabStop16 = new TabStop() { Val = TabStopValues.Number, Position = 1284 };

            tabs13.Append(tabStop16);
            Indentation indentation84 = new Indentation() { Start = "1284", Hanging = "720" };

            previousParagraphProperties46.Append(tabs13);
            previousParagraphProperties46.Append(indentation84);

            NumberingSymbolRunProperties numberingSymbolRunProperties17 = new NumberingSymbolRunProperties();
            RunFonts runFonts213 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties17.Append(runFonts213);

            level46.Append(startNumberingValue46);
            level46.Append(numberingFormat46);
            level46.Append(levelText46);
            level46.Append(levelJustification46);
            level46.Append(previousParagraphProperties46);
            level46.Append(numberingSymbolRunProperties17);

            Level level47 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue47 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat47 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText47 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification47 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties47 = new PreviousParagraphProperties();

            Tabs tabs14 = new Tabs();
            TabStop tabStop17 = new TabStop() { Val = TabStopValues.Number, Position = 1524 };

            tabs14.Append(tabStop17);
            Indentation indentation85 = new Indentation() { Start = "1524", Hanging = "480" };

            previousParagraphProperties47.Append(tabs14);
            previousParagraphProperties47.Append(indentation85);

            level47.Append(startNumberingValue47);
            level47.Append(numberingFormat47);
            level47.Append(levelText47);
            level47.Append(levelJustification47);
            level47.Append(previousParagraphProperties47);

            Level level48 = new Level() { LevelIndex = 2, TemplateCode = "0409001B", Tentative = true };
            StartNumberingValue startNumberingValue48 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat48 = new NumberingFormat() { Val = NumberFormatValues.LowerRoman };
            LevelText levelText48 = new LevelText() { Val = "%3." };
            LevelJustification levelJustification48 = new LevelJustification() { Val = LevelJustificationValues.Right };

            PreviousParagraphProperties previousParagraphProperties48 = new PreviousParagraphProperties();

            Tabs tabs15 = new Tabs();
            TabStop tabStop18 = new TabStop() { Val = TabStopValues.Number, Position = 2004 };

            tabs15.Append(tabStop18);
            Indentation indentation86 = new Indentation() { Start = "2004", Hanging = "480" };

            previousParagraphProperties48.Append(tabs15);
            previousParagraphProperties48.Append(indentation86);

            level48.Append(startNumberingValue48);
            level48.Append(numberingFormat48);
            level48.Append(levelText48);
            level48.Append(levelJustification48);
            level48.Append(previousParagraphProperties48);

            Level level49 = new Level() { LevelIndex = 3, TemplateCode = "0409000F", Tentative = true };
            StartNumberingValue startNumberingValue49 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat49 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText49 = new LevelText() { Val = "%4." };
            LevelJustification levelJustification49 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties49 = new PreviousParagraphProperties();

            Tabs tabs16 = new Tabs();
            TabStop tabStop19 = new TabStop() { Val = TabStopValues.Number, Position = 2484 };

            tabs16.Append(tabStop19);
            Indentation indentation87 = new Indentation() { Start = "2484", Hanging = "480" };

            previousParagraphProperties49.Append(tabs16);
            previousParagraphProperties49.Append(indentation87);

            level49.Append(startNumberingValue49);
            level49.Append(numberingFormat49);
            level49.Append(levelText49);
            level49.Append(levelJustification49);
            level49.Append(previousParagraphProperties49);

            Level level50 = new Level() { LevelIndex = 4, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue50 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat50 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText50 = new LevelText() { Val = "%5、" };
            LevelJustification levelJustification50 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties50 = new PreviousParagraphProperties();

            Tabs tabs17 = new Tabs();
            TabStop tabStop20 = new TabStop() { Val = TabStopValues.Number, Position = 2964 };

            tabs17.Append(tabStop20);
            Indentation indentation88 = new Indentation() { Start = "2964", Hanging = "480" };

            previousParagraphProperties50.Append(tabs17);
            previousParagraphProperties50.Append(indentation88);

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

            Tabs tabs18 = new Tabs();
            TabStop tabStop21 = new TabStop() { Val = TabStopValues.Number, Position = 3444 };

            tabs18.Append(tabStop21);
            Indentation indentation89 = new Indentation() { Start = "3444", Hanging = "480" };

            previousParagraphProperties51.Append(tabs18);
            previousParagraphProperties51.Append(indentation89);

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

            Tabs tabs19 = new Tabs();
            TabStop tabStop22 = new TabStop() { Val = TabStopValues.Number, Position = 3924 };

            tabs19.Append(tabStop22);
            Indentation indentation90 = new Indentation() { Start = "3924", Hanging = "480" };

            previousParagraphProperties52.Append(tabs19);
            previousParagraphProperties52.Append(indentation90);

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

            Tabs tabs20 = new Tabs();
            TabStop tabStop23 = new TabStop() { Val = TabStopValues.Number, Position = 4404 };

            tabs20.Append(tabStop23);
            Indentation indentation91 = new Indentation() { Start = "4404", Hanging = "480" };

            previousParagraphProperties53.Append(tabs20);
            previousParagraphProperties53.Append(indentation91);

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

            Tabs tabs21 = new Tabs();
            TabStop tabStop24 = new TabStop() { Val = TabStopValues.Number, Position = 4884 };

            tabs21.Append(tabStop24);
            Indentation indentation92 = new Indentation() { Start = "4884", Hanging = "480" };

            previousParagraphProperties54.Append(tabs21);
            previousParagraphProperties54.Append(indentation92);

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
            Nsid nsid7 = new Nsid() { Val = "41A04F1E" };
            MultiLevelType multiLevelType7 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode7 = new TemplateCode() { Val = "0024B006" };

            Level level55 = new Level() { LevelIndex = 0, TemplateCode = "55D8A5B4" };
            StartNumberingValue startNumberingValue55 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat55 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText55 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification55 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties55 = new PreviousParagraphProperties();
            Indentation indentation93 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties55.Append(indentation93);

            NumberingSymbolRunProperties numberingSymbolRunProperties18 = new NumberingSymbolRunProperties();
            RunFonts runFonts214 = new RunFonts() { Hint = FontTypeHintValues.Default };
            Color color129 = new Color() { Val = "000000" };

            numberingSymbolRunProperties18.Append(runFonts214);
            numberingSymbolRunProperties18.Append(color129);

            level55.Append(startNumberingValue55);
            level55.Append(numberingFormat55);
            level55.Append(levelText55);
            level55.Append(levelJustification55);
            level55.Append(previousParagraphProperties55);
            level55.Append(numberingSymbolRunProperties18);

            Level level56 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue56 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat56 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText56 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification56 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties56 = new PreviousParagraphProperties();
            Indentation indentation94 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties56.Append(indentation94);

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
            Indentation indentation95 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties57.Append(indentation95);

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
            Indentation indentation96 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties58.Append(indentation96);

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
            Indentation indentation97 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties59.Append(indentation97);

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
            Indentation indentation98 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties60.Append(indentation98);

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
            Indentation indentation99 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties61.Append(indentation99);

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
            Indentation indentation100 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties62.Append(indentation100);

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
            Indentation indentation101 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties63.Append(indentation101);

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
            Nsid nsid8 = new Nsid() { Val = "44176B93" };
            MultiLevelType multiLevelType8 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode8 = new TemplateCode() { Val = "54269EFA" };

            Level level64 = new Level() { LevelIndex = 0, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue64 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat64 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText64 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification64 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties64 = new PreviousParagraphProperties();
            Indentation indentation102 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties64.Append(indentation102);

            level64.Append(startNumberingValue64);
            level64.Append(numberingFormat64);
            level64.Append(levelText64);
            level64.Append(levelJustification64);
            level64.Append(previousParagraphProperties64);

            Level level65 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue65 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat65 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText65 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification65 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties65 = new PreviousParagraphProperties();
            Indentation indentation103 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties65.Append(indentation103);

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
            Indentation indentation104 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties66.Append(indentation104);

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
            Indentation indentation105 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties67.Append(indentation105);

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
            Indentation indentation106 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties68.Append(indentation106);

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
            Indentation indentation107 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties69.Append(indentation107);

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
            Indentation indentation108 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties70.Append(indentation108);

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
            Indentation indentation109 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties71.Append(indentation109);

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
            Indentation indentation110 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties72.Append(indentation110);

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
            Nsid nsid9 = new Nsid() { Val = "49EC268A" };
            MultiLevelType multiLevelType9 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode9 = new TemplateCode() { Val = "927E4F54" };

            Level level73 = new Level() { LevelIndex = 0, TemplateCode = "A6F0E89E" };
            StartNumberingValue startNumberingValue73 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat73 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText73 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification73 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties73 = new PreviousParagraphProperties();
            Indentation indentation111 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties73.Append(indentation111);

            NumberingSymbolRunProperties numberingSymbolRunProperties19 = new NumberingSymbolRunProperties();
            RunFonts runFonts215 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties19.Append(runFonts215);

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
            Indentation indentation112 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties74.Append(indentation112);

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
            Indentation indentation113 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties75.Append(indentation113);

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
            Indentation indentation114 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties76.Append(indentation114);

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
            Indentation indentation115 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties77.Append(indentation115);

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
            Indentation indentation116 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties78.Append(indentation116);

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
            Indentation indentation117 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties79.Append(indentation117);

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
            Indentation indentation118 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties80.Append(indentation118);

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
            Indentation indentation119 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties81.Append(indentation119);

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
            Nsid nsid10 = new Nsid() { Val = "64403189" };
            MultiLevelType multiLevelType10 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode10 = new TemplateCode() { Val = "0F50D422" };

            Level level82 = new Level() { LevelIndex = 0, TemplateCode = "0409000F" };
            StartNumberingValue startNumberingValue82 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat82 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText82 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification82 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties82 = new PreviousParagraphProperties();
            Indentation indentation120 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties82.Append(indentation120);

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
            Indentation indentation121 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties83.Append(indentation121);

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
            Indentation indentation122 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties84.Append(indentation122);

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
            Indentation indentation123 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties85.Append(indentation123);

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
            Indentation indentation124 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties86.Append(indentation124);

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
            Indentation indentation125 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties87.Append(indentation125);

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
            Indentation indentation126 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties88.Append(indentation126);

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
            Indentation indentation127 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties89.Append(indentation127);

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
            Indentation indentation128 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties90.Append(indentation128);

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
            Nsid nsid11 = new Nsid() { Val = "67E86D56" };
            MultiLevelType multiLevelType11 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode11 = new TemplateCode() { Val = "B7ACF502" };

            Level level91 = new Level() { LevelIndex = 0, TemplateCode = "04090015" };
            StartNumberingValue startNumberingValue91 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat91 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText91 = new LevelText() { Val = "%1、" };
            LevelJustification levelJustification91 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties91 = new PreviousParagraphProperties();
            Indentation indentation129 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties91.Append(indentation129);

            level91.Append(startNumberingValue91);
            level91.Append(numberingFormat91);
            level91.Append(levelText91);
            level91.Append(levelJustification91);
            level91.Append(previousParagraphProperties91);

            Level level92 = new Level() { LevelIndex = 1, TemplateCode = "04090019" };
            StartNumberingValue startNumberingValue92 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat92 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText92 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification92 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties92 = new PreviousParagraphProperties();
            Indentation indentation130 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties92.Append(indentation130);

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
            Indentation indentation131 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties93.Append(indentation131);

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
            Indentation indentation132 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties94.Append(indentation132);

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
            Indentation indentation133 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties95.Append(indentation133);

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
            Indentation indentation134 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties96.Append(indentation134);

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
            Indentation indentation135 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties97.Append(indentation135);

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
            Indentation indentation136 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties98.Append(indentation136);

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
            Indentation indentation137 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties99.Append(indentation137);

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
            Nsid nsid12 = new Nsid() { Val = "72347E41" };
            MultiLevelType multiLevelType12 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode12 = new TemplateCode() { Val = "D58AC6F4" };

            Level level100 = new Level() { LevelIndex = 0, TemplateCode = "1046A65E" };
            StartNumberingValue startNumberingValue100 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat100 = new NumberingFormat() { Val = NumberFormatValues.TaiwaneseCountingThousand };
            LevelText levelText100 = new LevelText() { Val = "(%1)" };
            LevelJustification levelJustification100 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties100 = new PreviousParagraphProperties();

            Tabs tabs22 = new Tabs();
            TabStop tabStop25 = new TabStop() { Val = TabStopValues.Number, Position = 1284 };

            tabs22.Append(tabStop25);
            Indentation indentation138 = new Indentation() { Start = "1284", Hanging = "720" };

            previousParagraphProperties100.Append(tabs22);
            previousParagraphProperties100.Append(indentation138);

            NumberingSymbolRunProperties numberingSymbolRunProperties20 = new NumberingSymbolRunProperties();
            RunFonts runFonts216 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            numberingSymbolRunProperties20.Append(runFonts216);

            level100.Append(startNumberingValue100);
            level100.Append(numberingFormat100);
            level100.Append(levelText100);
            level100.Append(levelJustification100);
            level100.Append(previousParagraphProperties100);
            level100.Append(numberingSymbolRunProperties20);

            Level level101 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue101 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat101 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText101 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification101 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties101 = new PreviousParagraphProperties();

            Tabs tabs23 = new Tabs();
            TabStop tabStop26 = new TabStop() { Val = TabStopValues.Number, Position = 1524 };

            tabs23.Append(tabStop26);
            Indentation indentation139 = new Indentation() { Start = "1524", Hanging = "480" };

            previousParagraphProperties101.Append(tabs23);
            previousParagraphProperties101.Append(indentation139);

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

            Tabs tabs24 = new Tabs();
            TabStop tabStop27 = new TabStop() { Val = TabStopValues.Number, Position = 2004 };

            tabs24.Append(tabStop27);
            Indentation indentation140 = new Indentation() { Start = "2004", Hanging = "480" };

            previousParagraphProperties102.Append(tabs24);
            previousParagraphProperties102.Append(indentation140);

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

            Tabs tabs25 = new Tabs();
            TabStop tabStop28 = new TabStop() { Val = TabStopValues.Number, Position = 2484 };

            tabs25.Append(tabStop28);
            Indentation indentation141 = new Indentation() { Start = "2484", Hanging = "480" };

            previousParagraphProperties103.Append(tabs25);
            previousParagraphProperties103.Append(indentation141);

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

            Tabs tabs26 = new Tabs();
            TabStop tabStop29 = new TabStop() { Val = TabStopValues.Number, Position = 2964 };

            tabs26.Append(tabStop29);
            Indentation indentation142 = new Indentation() { Start = "2964", Hanging = "480" };

            previousParagraphProperties104.Append(tabs26);
            previousParagraphProperties104.Append(indentation142);

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

            Tabs tabs27 = new Tabs();
            TabStop tabStop30 = new TabStop() { Val = TabStopValues.Number, Position = 3444 };

            tabs27.Append(tabStop30);
            Indentation indentation143 = new Indentation() { Start = "3444", Hanging = "480" };

            previousParagraphProperties105.Append(tabs27);
            previousParagraphProperties105.Append(indentation143);

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

            Tabs tabs28 = new Tabs();
            TabStop tabStop31 = new TabStop() { Val = TabStopValues.Number, Position = 3924 };

            tabs28.Append(tabStop31);
            Indentation indentation144 = new Indentation() { Start = "3924", Hanging = "480" };

            previousParagraphProperties106.Append(tabs28);
            previousParagraphProperties106.Append(indentation144);

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

            Tabs tabs29 = new Tabs();
            TabStop tabStop32 = new TabStop() { Val = TabStopValues.Number, Position = 4404 };

            tabs29.Append(tabStop32);
            Indentation indentation145 = new Indentation() { Start = "4404", Hanging = "480" };

            previousParagraphProperties107.Append(tabs29);
            previousParagraphProperties107.Append(indentation145);

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

            Tabs tabs30 = new Tabs();
            TabStop tabStop33 = new TabStop() { Val = TabStopValues.Number, Position = 4884 };

            tabs30.Append(tabStop33);
            Indentation indentation146 = new Indentation() { Start = "4884", Hanging = "480" };

            previousParagraphProperties108.Append(tabs30);
            previousParagraphProperties108.Append(indentation146);

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
            Nsid nsid13 = new Nsid() { Val = "72F72D62" };
            MultiLevelType multiLevelType13 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode13 = new TemplateCode() { Val = "C03079DE" };

            Level level109 = new Level() { LevelIndex = 0, TemplateCode = "F4EEFD4A" };
            StartNumberingValue startNumberingValue109 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat109 = new NumberingFormat() { Val = NumberFormatValues.Decimal };
            LevelText levelText109 = new LevelText() { Val = "%1." };
            LevelJustification levelJustification109 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties109 = new PreviousParagraphProperties();
            Indentation indentation147 = new Indentation() { Start = "360", Hanging = "360" };

            previousParagraphProperties109.Append(indentation147);

            NumberingSymbolRunProperties numberingSymbolRunProperties21 = new NumberingSymbolRunProperties();
            RunFonts runFonts217 = new RunFonts() { Hint = FontTypeHintValues.Default };

            numberingSymbolRunProperties21.Append(runFonts217);

            level109.Append(startNumberingValue109);
            level109.Append(numberingFormat109);
            level109.Append(levelText109);
            level109.Append(levelJustification109);
            level109.Append(previousParagraphProperties109);
            level109.Append(numberingSymbolRunProperties21);

            Level level110 = new Level() { LevelIndex = 1, TemplateCode = "04090019", Tentative = true };
            StartNumberingValue startNumberingValue110 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat110 = new NumberingFormat() { Val = NumberFormatValues.IdeographTraditional };
            LevelText levelText110 = new LevelText() { Val = "%2、" };
            LevelJustification levelJustification110 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties110 = new PreviousParagraphProperties();
            Indentation indentation148 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties110.Append(indentation148);

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
            Indentation indentation149 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties111.Append(indentation149);

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
            Indentation indentation150 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties112.Append(indentation150);

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
            Indentation indentation151 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties113.Append(indentation151);

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
            Indentation indentation152 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties114.Append(indentation152);

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
            Indentation indentation153 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties115.Append(indentation153);

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
            Indentation indentation154 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties116.Append(indentation154);

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
            Indentation indentation155 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties117.Append(indentation155);

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
            Nsid nsid14 = new Nsid() { Val = "7BAF4830" };
            MultiLevelType multiLevelType14 = new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel };
            TemplateCode templateCode14 = new TemplateCode() { Val = "4C9690DC" };

            Level level118 = new Level() { LevelIndex = 0, TemplateCode = "6EF420B6" };
            StartNumberingValue startNumberingValue118 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat118 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText118 = new LevelText() { Val = "Ÿ" };
            LevelJustification levelJustification118 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties118 = new PreviousParagraphProperties();
            Indentation indentation156 = new Indentation() { Start = "480", Hanging = "480" };

            previousParagraphProperties118.Append(indentation156);

            NumberingSymbolRunProperties numberingSymbolRunProperties22 = new NumberingSymbolRunProperties();
            RunFonts runFonts218 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties22.Append(runFonts218);

            level118.Append(startNumberingValue118);
            level118.Append(numberingFormat118);
            level118.Append(levelText118);
            level118.Append(levelJustification118);
            level118.Append(previousParagraphProperties118);
            level118.Append(numberingSymbolRunProperties22);

            Level level119 = new Level() { LevelIndex = 1, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue119 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat119 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText119 = new LevelText() { Val = "n" };
            LevelJustification levelJustification119 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties119 = new PreviousParagraphProperties();
            Indentation indentation157 = new Indentation() { Start = "960", Hanging = "480" };

            previousParagraphProperties119.Append(indentation157);

            NumberingSymbolRunProperties numberingSymbolRunProperties23 = new NumberingSymbolRunProperties();
            RunFonts runFonts219 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties23.Append(runFonts219);

            level119.Append(startNumberingValue119);
            level119.Append(numberingFormat119);
            level119.Append(levelText119);
            level119.Append(levelJustification119);
            level119.Append(previousParagraphProperties119);
            level119.Append(numberingSymbolRunProperties23);

            Level level120 = new Level() { LevelIndex = 2, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue120 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat120 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText120 = new LevelText() { Val = "u" };
            LevelJustification levelJustification120 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties120 = new PreviousParagraphProperties();
            Indentation indentation158 = new Indentation() { Start = "1440", Hanging = "480" };

            previousParagraphProperties120.Append(indentation158);

            NumberingSymbolRunProperties numberingSymbolRunProperties24 = new NumberingSymbolRunProperties();
            RunFonts runFonts220 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties24.Append(runFonts220);

            level120.Append(startNumberingValue120);
            level120.Append(numberingFormat120);
            level120.Append(levelText120);
            level120.Append(levelJustification120);
            level120.Append(previousParagraphProperties120);
            level120.Append(numberingSymbolRunProperties24);

            Level level121 = new Level() { LevelIndex = 3, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue121 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat121 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText121 = new LevelText() { Val = "l" };
            LevelJustification levelJustification121 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties121 = new PreviousParagraphProperties();
            Indentation indentation159 = new Indentation() { Start = "1920", Hanging = "480" };

            previousParagraphProperties121.Append(indentation159);

            NumberingSymbolRunProperties numberingSymbolRunProperties25 = new NumberingSymbolRunProperties();
            RunFonts runFonts221 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties25.Append(runFonts221);

            level121.Append(startNumberingValue121);
            level121.Append(numberingFormat121);
            level121.Append(levelText121);
            level121.Append(levelJustification121);
            level121.Append(previousParagraphProperties121);
            level121.Append(numberingSymbolRunProperties25);

            Level level122 = new Level() { LevelIndex = 4, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue122 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat122 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText122 = new LevelText() { Val = "n" };
            LevelJustification levelJustification122 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties122 = new PreviousParagraphProperties();
            Indentation indentation160 = new Indentation() { Start = "2400", Hanging = "480" };

            previousParagraphProperties122.Append(indentation160);

            NumberingSymbolRunProperties numberingSymbolRunProperties26 = new NumberingSymbolRunProperties();
            RunFonts runFonts222 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties26.Append(runFonts222);

            level122.Append(startNumberingValue122);
            level122.Append(numberingFormat122);
            level122.Append(levelText122);
            level122.Append(levelJustification122);
            level122.Append(previousParagraphProperties122);
            level122.Append(numberingSymbolRunProperties26);

            Level level123 = new Level() { LevelIndex = 5, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue123 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat123 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText123 = new LevelText() { Val = "u" };
            LevelJustification levelJustification123 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties123 = new PreviousParagraphProperties();
            Indentation indentation161 = new Indentation() { Start = "2880", Hanging = "480" };

            previousParagraphProperties123.Append(indentation161);

            NumberingSymbolRunProperties numberingSymbolRunProperties27 = new NumberingSymbolRunProperties();
            RunFonts runFonts223 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties27.Append(runFonts223);

            level123.Append(startNumberingValue123);
            level123.Append(numberingFormat123);
            level123.Append(levelText123);
            level123.Append(levelJustification123);
            level123.Append(previousParagraphProperties123);
            level123.Append(numberingSymbolRunProperties27);

            Level level124 = new Level() { LevelIndex = 6, TemplateCode = "04090001", Tentative = true };
            StartNumberingValue startNumberingValue124 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat124 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText124 = new LevelText() { Val = "l" };
            LevelJustification levelJustification124 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties124 = new PreviousParagraphProperties();
            Indentation indentation162 = new Indentation() { Start = "3360", Hanging = "480" };

            previousParagraphProperties124.Append(indentation162);

            NumberingSymbolRunProperties numberingSymbolRunProperties28 = new NumberingSymbolRunProperties();
            RunFonts runFonts224 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties28.Append(runFonts224);

            level124.Append(startNumberingValue124);
            level124.Append(numberingFormat124);
            level124.Append(levelText124);
            level124.Append(levelJustification124);
            level124.Append(previousParagraphProperties124);
            level124.Append(numberingSymbolRunProperties28);

            Level level125 = new Level() { LevelIndex = 7, TemplateCode = "04090003", Tentative = true };
            StartNumberingValue startNumberingValue125 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat125 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText125 = new LevelText() { Val = "n" };
            LevelJustification levelJustification125 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties125 = new PreviousParagraphProperties();
            Indentation indentation163 = new Indentation() { Start = "3840", Hanging = "480" };

            previousParagraphProperties125.Append(indentation163);

            NumberingSymbolRunProperties numberingSymbolRunProperties29 = new NumberingSymbolRunProperties();
            RunFonts runFonts225 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties29.Append(runFonts225);

            level125.Append(startNumberingValue125);
            level125.Append(numberingFormat125);
            level125.Append(levelText125);
            level125.Append(levelJustification125);
            level125.Append(previousParagraphProperties125);
            level125.Append(numberingSymbolRunProperties29);

            Level level126 = new Level() { LevelIndex = 8, TemplateCode = "04090005", Tentative = true };
            StartNumberingValue startNumberingValue126 = new StartNumberingValue() { Val = 1 };
            NumberingFormat numberingFormat126 = new NumberingFormat() { Val = NumberFormatValues.Bullet };
            LevelText levelText126 = new LevelText() { Val = "u" };
            LevelJustification levelJustification126 = new LevelJustification() { Val = LevelJustificationValues.Left };

            PreviousParagraphProperties previousParagraphProperties126 = new PreviousParagraphProperties();
            Indentation indentation164 = new Indentation() { Start = "4320", Hanging = "480" };

            previousParagraphProperties126.Append(indentation164);

            NumberingSymbolRunProperties numberingSymbolRunProperties30 = new NumberingSymbolRunProperties();
            RunFonts runFonts226 = new RunFonts() { Hint = FontTypeHintValues.Default, Ascii = "Wingdings", HighAnsi = "Wingdings" };

            numberingSymbolRunProperties30.Append(runFonts226);

            level126.Append(startNumberingValue126);
            level126.Append(numberingFormat126);
            level126.Append(levelText126);
            level126.Append(levelJustification126);
            level126.Append(previousParagraphProperties126);
            level126.Append(numberingSymbolRunProperties30);

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

            NumberingInstance numberingInstance1 = new NumberingInstance() { NumberID = 1 };
            numberingInstance1.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1144665447"));
            AbstractNumId abstractNumId1 = new AbstractNumId() { Val = 5 };

            numberingInstance1.Append(abstractNumId1);

            NumberingInstance numberingInstance2 = new NumberingInstance() { NumberID = 2 };
            numberingInstance2.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "2142259759"));
            AbstractNumId abstractNumId2 = new AbstractNumId() { Val = 9 };

            numberingInstance2.Append(abstractNumId2);

            NumberingInstance numberingInstance3 = new NumberingInstance() { NumberID = 3 };
            numberingInstance3.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "550772816"));
            AbstractNumId abstractNumId3 = new AbstractNumId() { Val = 4 };

            numberingInstance3.Append(abstractNumId3);

            NumberingInstance numberingInstance4 = new NumberingInstance() { NumberID = 4 };
            numberingInstance4.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1403285214"));
            AbstractNumId abstractNumId4 = new AbstractNumId() { Val = 1 };

            numberingInstance4.Append(abstractNumId4);

            NumberingInstance numberingInstance5 = new NumberingInstance() { NumberID = 5 };
            numberingInstance5.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "2009405194"));
            AbstractNumId abstractNumId5 = new AbstractNumId() { Val = 8 };

            numberingInstance5.Append(abstractNumId5);

            NumberingInstance numberingInstance6 = new NumberingInstance() { NumberID = 6 };
            numberingInstance6.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "508103745"));
            AbstractNumId abstractNumId6 = new AbstractNumId() { Val = 12 };

            numberingInstance6.Append(abstractNumId6);

            NumberingInstance numberingInstance7 = new NumberingInstance() { NumberID = 7 };
            numberingInstance7.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1312754357"));
            AbstractNumId abstractNumId7 = new AbstractNumId() { Val = 13 };

            numberingInstance7.Append(abstractNumId7);

            NumberingInstance numberingInstance8 = new NumberingInstance() { NumberID = 8 };
            numberingInstance8.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1708337356"));
            AbstractNumId abstractNumId8 = new AbstractNumId() { Val = 3 };

            numberingInstance8.Append(abstractNumId8);

            NumberingInstance numberingInstance9 = new NumberingInstance() { NumberID = 9 };
            numberingInstance9.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1693650854"));
            AbstractNumId abstractNumId9 = new AbstractNumId() { Val = 11 };

            numberingInstance9.Append(abstractNumId9);

            NumberingInstance numberingInstance10 = new NumberingInstance() { NumberID = 10 };
            numberingInstance10.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "245963574"));
            AbstractNumId abstractNumId10 = new AbstractNumId() { Val = 10 };

            numberingInstance10.Append(abstractNumId10);

            NumberingInstance numberingInstance11 = new NumberingInstance() { NumberID = 11 };
            numberingInstance11.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "579680535"));
            AbstractNumId abstractNumId11 = new AbstractNumId() { Val = 0 };

            numberingInstance11.Append(abstractNumId11);

            NumberingInstance numberingInstance12 = new NumberingInstance() { NumberID = 12 };
            numberingInstance12.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1511679809"));
            AbstractNumId abstractNumId12 = new AbstractNumId() { Val = 2 };

            numberingInstance12.Append(abstractNumId12);

            NumberingInstance numberingInstance13 = new NumberingInstance() { NumberID = 13 };
            numberingInstance13.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "1844054497"));
            AbstractNumId abstractNumId13 = new AbstractNumId() { Val = 7 };

            numberingInstance13.Append(abstractNumId13);

            NumberingInstance numberingInstance14 = new NumberingInstance() { NumberID = 14 };
            numberingInstance14.SetAttribute(new OpenXmlAttribute("w16cid", "durableId", "http://schemas.microsoft.com/office/word/2016/wordml/cid", "818155963"));
            AbstractNumId abstractNumId14 = new AbstractNumId() { Val = 6 };

            numberingInstance14.Append(abstractNumId14);

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

            Paragraph paragraph84 = new Paragraph() { RsidParagraphAddition = "00DC71B6", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00DC71B6", ParagraphId = "2858168F", TextId = "77777777" };

            Run run195 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run195.Append(separatorMark1);

            paragraph84.Append(run195);

            endnote1.Append(paragraph84);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph85 = new Paragraph() { RsidParagraphAddition = "00DC71B6", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00DC71B6", ParagraphId = "27BC1FAB", TextId = "77777777" };

            Run run196 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run196.Append(continuationSeparatorMark1);

            paragraph85.Append(run196);

            endnote2.Append(paragraph85);

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
            Ds.DataStoreItem dataStoreItem4 = new Ds.DataStoreItem() { ItemId = "{160B8829-BD19-448B-8CC7-59FD88F579EA}" };
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

            Paragraph paragraph86 = new Paragraph() { RsidParagraphAddition = "00DC71B6", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00DC71B6", ParagraphId = "178DC6C3", TextId = "77777777" };

            Run run197 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run197.Append(separatorMark2);

            paragraph86.Append(run197);

            footnote1.Append(paragraph86);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph87 = new Paragraph() { RsidParagraphAddition = "00DC71B6", RsidParagraphProperties = "003C3877", RsidRunAdditionDefault = "00DC71B6", ParagraphId = "14EE71B5", TextId = "77777777" };

            Run run198 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run198.Append(continuationSeparatorMark2);

            paragraph87.Append(run198);

            footnote2.Append(paragraph87);

            Footnote footnote3 = new Footnote() { Id = 1 };

            Paragraph paragraph88 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00D43514", RsidParagraphProperties = "001C3E6D", RsidRunAdditionDefault = "006C3779", ParagraphId = "1A94DB08", TextId = "77777777" };

            ParagraphProperties paragraphProperties81 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines35 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation165 = new Indentation() { Start = "64", Hanging = "64", HangingChars = 40 };

            ParagraphMarkRunProperties paragraphMarkRunProperties74 = new ParagraphMarkRunProperties();
            FontSize fontSize237 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript90 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties74.Append(fontSize237);
            paragraphMarkRunProperties74.Append(fontSizeComplexScript90);

            paragraphProperties81.Append(spacingBetweenLines35);
            paragraphProperties81.Append(indentation165);
            paragraphProperties81.Append(paragraphMarkRunProperties74);

            Run run199 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties191 = new RunProperties();
            RunStyle runStyle11 = new RunStyle() { Val = "a9" };
            FontSize fontSize238 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript91 = new FontSizeComplexScript() { Val = "16" };

            runProperties191.Append(runStyle11);
            runProperties191.Append(fontSize238);
            runProperties191.Append(fontSizeComplexScript91);
            FootnoteReferenceMark footnoteReferenceMark1 = new FootnoteReferenceMark();

            run199.Append(runProperties191);
            run199.Append(footnoteReferenceMark1);

            Run run200 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties192 = new RunProperties();
            FontSize fontSize239 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript92 = new FontSizeComplexScript() { Val = "16" };

            runProperties192.Append(fontSize239);
            runProperties192.Append(fontSizeComplexScript92);
            Text text183 = new Text();
            text183.Text = "「";

            run200.Append(runProperties192);
            run200.Append(text183);

            Run run201 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties193 = new RunProperties();
            FontSize fontSize240 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript93 = new FontSizeComplexScript() { Val = "16" };

            runProperties193.Append(fontSize240);
            runProperties193.Append(fontSizeComplexScript93);
            Text text184 = new Text();
            text184.Text = "外部法規之變化」";

            run201.Append(runProperties193);
            run201.Append(text184);

            Run run202 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties194 = new RunProperties();
            RunFonts runFonts227 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize241 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript94 = new FontSizeComplexScript() { Val = "16" };

            runProperties194.Append(runFonts227);
            runProperties194.Append(fontSize241);
            runProperties194.Append(fontSizeComplexScript94);
            Text text185 = new Text();
            text185.Text = "應";

            run202.Append(runProperties194);
            run202.Append(text185);

            Run run203 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties195 = new RunProperties();
            FontSize fontSize242 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript95 = new FontSizeComplexScript() { Val = "16" };

            runProperties195.Append(fontSize242);
            runProperties195.Append(fontSizeComplexScript95);
            Text text186 = new Text();
            text186.Text = "揭露前次查核後";

            run203.Append(runProperties195);
            run203.Append(text186);

            Run run204 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties196 = new RunProperties();
            RunFonts runFonts228 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize243 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript96 = new FontSizeComplexScript() { Val = "16" };

            runProperties196.Append(runFonts228);
            runProperties196.Append(fontSize243);
            runProperties196.Append(fontSizeComplexScript96);
            Text text187 = new Text();
            text187.Text = "迄基準日與受檢單位業務相關的重要";

            run204.Append(runProperties196);
            run204.Append(text187);

            Run run205 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties197 = new RunProperties();
            FontSize fontSize244 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript97 = new FontSizeComplexScript() { Val = "16" };

            runProperties197.Append(fontSize244);
            runProperties197.Append(fontSizeComplexScript97);
            Text text188 = new Text();
            text188.Text = "外部法規";

            run205.Append(runProperties197);
            run205.Append(text188);

            Run run206 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties198 = new RunProperties();
            RunFonts runFonts229 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize245 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript98 = new FontSizeComplexScript() { Val = "16" };

            runProperties198.Append(runFonts229);
            runProperties198.Append(fontSize245);
            runProperties198.Append(fontSizeComplexScript98);
            Text text189 = new Text();
            text189.Text = "增修訂之情形";

            run206.Append(runProperties198);
            run206.Append(text189);

            Run run207 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties199 = new RunProperties();
            FontSize fontSize246 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript99 = new FontSizeComplexScript() { Val = "16" };

            runProperties199.Append(fontSize246);
            runProperties199.Append(fontSizeComplexScript99);
            Text text190 = new Text();
            text190.Text = "，";

            run207.Append(runProperties199);
            run207.Append(text190);

            Run run208 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties200 = new RunProperties();
            RunFonts runFonts230 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize247 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript100 = new FontSizeComplexScript() { Val = "16" };

            runProperties200.Append(runFonts230);
            runProperties200.Append(fontSize247);
            runProperties200.Append(fontSizeComplexScript100);
            Text text191 = new Text();
            text191.Text = "內容原則";

            run208.Append(runProperties200);
            run208.Append(text191);
            ProofError proofError9 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run209 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties201 = new RunProperties();
            FontSize fontSize248 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript101 = new FontSizeComplexScript() { Val = "16" };

            runProperties201.Append(fontSize248);
            runProperties201.Append(fontSizeComplexScript101);
            Text text192 = new Text();
            text192.Text = "採";

            run209.Append(runProperties201);
            run209.Append(text192);
            ProofError proofError10 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run210 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties202 = new RunProperties();
            FontSize fontSize249 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript102 = new FontSizeComplexScript() { Val = "16" };

            runProperties202.Append(fontSize249);
            runProperties202.Append(fontSizeComplexScript102);
            Text text193 = new Text();
            text193.Text = "逐條揭露";

            run210.Append(runProperties202);
            run210.Append(text193);

            Run run211 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties203 = new RunProperties();
            RunFonts runFonts231 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize250 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript103 = new FontSizeComplexScript() { Val = "16" };

            runProperties203.Append(runFonts231);
            runProperties203.Append(fontSize250);
            runProperties203.Append(fontSizeComplexScript103);
            Text text194 = new Text();
            text194.Text = "，惟各科";

            run211.Append(runProperties203);
            run211.Append(text194);

            Run run212 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties204 = new RunProperties();
            FontSize fontSize251 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript104 = new FontSizeComplexScript() { Val = "16" };

            runProperties204.Append(fontSize251);
            runProperties204.Append(fontSizeComplexScript104);
            Text text195 = new Text();
            text195.Text = "實務作業";

            run212.Append(runProperties204);
            run212.Append(text195);

            Run run213 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties205 = new RunProperties();
            RunFonts runFonts232 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize252 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript105 = new FontSizeComplexScript() { Val = "16" };

            runProperties205.Append(runFonts232);
            runProperties205.Append(fontSize252);
            runProperties205.Append(fontSizeComplexScript105);
            Text text196 = new Text();
            text196.Text = "如為";

            run213.Append(runProperties205);
            run213.Append(text196);

            Run run214 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties206 = new RunProperties();
            FontSize fontSize253 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript106 = new FontSizeComplexScript() { Val = "16" };

            runProperties206.Append(fontSize253);
            runProperties206.Append(fontSizeComplexScript106);
            Text text197 = new Text();
            text197.Text = "日常";

            run214.Append(runProperties206);
            run214.Append(text197);

            Run run215 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties207 = new RunProperties();
            RunFonts runFonts233 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize254 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript107 = new FontSizeComplexScript() { Val = "16" };

            runProperties207.Append(runFonts233);
            runProperties207.Append(fontSize254);
            runProperties207.Append(fontSizeComplexScript107);
            Text text198 = new Text();
            text198.Text = "轄區資料夾維護";

            run215.Append(runProperties207);
            run215.Append(text198);

            Run run216 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties208 = new RunProperties();
            RunFonts runFonts234 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize255 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript108 = new FontSizeComplexScript() { Val = "16" };

            runProperties208.Append(runFonts234);
            runProperties208.Append(fontSize255);
            runProperties208.Append(fontSizeComplexScript108);
            Text text199 = new Text();
            text199.Text = "/";

            run216.Append(runProperties208);
            run216.Append(text199);

            Run run217 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties209 = new RunProperties();
            RunFonts runFonts235 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize256 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript109 = new FontSizeComplexScript() { Val = "16" };

            runProperties209.Append(runFonts235);
            runProperties209.Append(fontSize256);
            runProperties209.Append(fontSizeComplexScript109);
            Text text200 = new Text();
            text200.Text = "即時";

            run217.Append(runProperties209);
            run217.Append(text200);

            Run run218 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties210 = new RunProperties();
            FontSize fontSize257 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript110 = new FontSizeComplexScript() { Val = "16" };

            runProperties210.Append(fontSize257);
            runProperties210.Append(fontSizeComplexScript110);
            Text text201 = new Text();
            text201.Text = "修訂查核題庫";

            run218.Append(runProperties210);
            run218.Append(text201);

            Run run219 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties211 = new RunProperties();
            RunFonts runFonts236 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize258 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript111 = new FontSizeComplexScript() { Val = "16" };

            runProperties211.Append(runFonts236);
            runProperties211.Append(fontSize258);
            runProperties211.Append(fontSizeComplexScript111);
            Text text202 = new Text();
            text202.Text = "時，得以註明法規異動之索引資料來源代替";

            run219.Append(runProperties211);
            run219.Append(text202);
            ProofError proofError11 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run220 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties212 = new RunProperties();
            FontSize fontSize259 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript112 = new FontSizeComplexScript() { Val = "16" };

            runProperties212.Append(fontSize259);
            runProperties212.Append(fontSizeComplexScript112);
            Text text203 = new Text();
            text203.Text = "（";

            run220.Append(runProperties212);
            run220.Append(text203);
            ProofError proofError12 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run221 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties213 = new RunProperties();
            RunFonts runFonts237 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize260 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript113 = new FontSizeComplexScript() { Val = "16" };

            runProperties213.Append(runFonts237);
            runProperties213.Append(fontSize260);
            runProperties213.Append(fontSizeComplexScript113);
            Text text204 = new Text();
            text204.Text = "例如：參";

            run221.Append(runProperties213);
            run221.Append(text204);
            ProofError proofError13 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run222 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties214 = new RunProperties();
            RunFonts runFonts238 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize261 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript114 = new FontSizeComplexScript() { Val = "16" };

            runProperties214.Append(runFonts238);
            runProperties214.Append(fontSize261);
            runProperties214.Append(fontSizeComplexScript114);
            Text text205 = new Text();
            text205.Text = "xxxx";

            run222.Append(runProperties214);
            run222.Append(text205);
            ProofError proofError14 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run223 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties215 = new RunProperties();
            RunFonts runFonts239 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize262 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript115 = new FontSizeComplexScript() { Val = "16" };

            runProperties215.Append(runFonts239);
            runProperties215.Append(fontSize262);
            runProperties215.Append(fontSizeComplexScript115);
            Text text206 = new Text();
            text206.Text = "題庫、轄區";

            run223.Append(runProperties215);
            run223.Append(text206);
            ProofError proofError15 = new ProofError() { Type = ProofingErrorValues.SpellStart };

            Run run224 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties216 = new RunProperties();
            RunFonts runFonts240 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize263 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript116 = new FontSizeComplexScript() { Val = "16" };

            runProperties216.Append(runFonts240);
            runProperties216.Append(fontSize263);
            runProperties216.Append(fontSizeComplexScript116);
            Text text207 = new Text();
            text207.Text = "xxxx";

            run224.Append(runProperties216);
            run224.Append(text207);
            ProofError proofError16 = new ProofError() { Type = ProofingErrorValues.SpellEnd };

            Run run225 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties217 = new RunProperties();
            FontSize fontSize264 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript117 = new FontSizeComplexScript() { Val = "16" };

            runProperties217.Append(fontSize264);
            runProperties217.Append(fontSizeComplexScript117);
            Text text208 = new Text();
            text208.Text = "資料";

            run225.Append(runProperties217);
            run225.Append(text208);

            Run run226 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties218 = new RunProperties();
            RunFonts runFonts241 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize265 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript118 = new FontSizeComplexScript() { Val = "16" };

            runProperties218.Append(runFonts241);
            runProperties218.Append(fontSize265);
            runProperties218.Append(fontSizeComplexScript118);
            Text text209 = new Text();
            text209.Text = "夾";

            run226.Append(runProperties218);
            run226.Append(text209);
            ProofError proofError17 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run227 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties219 = new RunProperties();
            FontSize fontSize266 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript119 = new FontSizeComplexScript() { Val = "16" };

            runProperties219.Append(fontSize266);
            runProperties219.Append(fontSizeComplexScript119);
            Text text210 = new Text();
            text210.Text = "）";

            run227.Append(runProperties219);
            run227.Append(text210);
            ProofError proofError18 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run228 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties220 = new RunProperties();
            RunFonts runFonts242 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize267 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript120 = new FontSizeComplexScript() { Val = "16" };

            runProperties220.Append(runFonts242);
            runProperties220.Append(fontSize267);
            runProperties220.Append(fontSizeComplexScript120);
            Text text211 = new Text();
            text211.Text = "，";

            run228.Append(runProperties220);
            run228.Append(text211);
            ProofError proofError19 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run229 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties221 = new RunProperties();
            RunFonts runFonts243 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize268 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript121 = new FontSizeComplexScript() { Val = "16" };

            runProperties221.Append(runFonts243);
            runProperties221.Append(fontSize268);
            runProperties221.Append(fontSizeComplexScript121);
            Text text212 = new Text();
            text212.Text = "俾";

            run229.Append(runProperties221);
            run229.Append(text212);
            ProofError proofError20 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run230 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00F03C81" };

            RunProperties runProperties222 = new RunProperties();
            RunFonts runFonts244 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize269 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript122 = new FontSizeComplexScript() { Val = "16" };

            runProperties222.Append(runFonts244);
            runProperties222.Append(fontSize269);
            runProperties222.Append(fontSizeComplexScript122);
            Text text213 = new Text();
            text213.Text = "利查核人員清楚知悉法規變化情形。";

            run230.Append(runProperties222);
            run230.Append(text213);

            paragraph88.Append(paragraphProperties81);
            paragraph88.Append(run199);
            paragraph88.Append(run200);
            paragraph88.Append(run201);
            paragraph88.Append(run202);
            paragraph88.Append(run203);
            paragraph88.Append(run204);
            paragraph88.Append(run205);
            paragraph88.Append(run206);
            paragraph88.Append(run207);
            paragraph88.Append(run208);
            paragraph88.Append(proofError9);
            paragraph88.Append(run209);
            paragraph88.Append(proofError10);
            paragraph88.Append(run210);
            paragraph88.Append(run211);
            paragraph88.Append(run212);
            paragraph88.Append(run213);
            paragraph88.Append(run214);
            paragraph88.Append(run215);
            paragraph88.Append(run216);
            paragraph88.Append(run217);
            paragraph88.Append(run218);
            paragraph88.Append(run219);
            paragraph88.Append(proofError11);
            paragraph88.Append(run220);
            paragraph88.Append(proofError12);
            paragraph88.Append(run221);
            paragraph88.Append(proofError13);
            paragraph88.Append(run222);
            paragraph88.Append(proofError14);
            paragraph88.Append(run223);
            paragraph88.Append(proofError15);
            paragraph88.Append(run224);
            paragraph88.Append(proofError16);
            paragraph88.Append(run225);
            paragraph88.Append(run226);
            paragraph88.Append(proofError17);
            paragraph88.Append(run227);
            paragraph88.Append(proofError18);
            paragraph88.Append(run228);
            paragraph88.Append(proofError19);
            paragraph88.Append(run229);
            paragraph88.Append(proofError20);
            paragraph88.Append(run230);

            footnote3.Append(paragraph88);

            Footnote footnote4 = new Footnote() { Id = 2 };

            Paragraph paragraph89 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00CF43E8", RsidRunAdditionDefault = "00CF43E8", ParagraphId = "5BA390CE", TextId = "77777777" };

            ParagraphProperties paragraphProperties82 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId23 = new ParagraphStyleId() { Val = "a7" };

            ParagraphMarkRunProperties paragraphMarkRunProperties75 = new ParagraphMarkRunProperties();
            FontSize fontSize270 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript123 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties75.Append(fontSize270);
            paragraphMarkRunProperties75.Append(fontSizeComplexScript123);

            paragraphProperties82.Append(paragraphStyleId23);
            paragraphProperties82.Append(paragraphMarkRunProperties75);

            Run run231 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties223 = new RunProperties();
            RunStyle runStyle12 = new RunStyle() { Val = "a9" };
            FontSize fontSize271 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript124 = new FontSizeComplexScript() { Val = "16" };

            runProperties223.Append(runStyle12);
            runProperties223.Append(fontSize271);
            runProperties223.Append(fontSizeComplexScript124);
            FootnoteReferenceMark footnoteReferenceMark2 = new FootnoteReferenceMark();

            run231.Append(runProperties223);
            run231.Append(footnoteReferenceMark2);

            Run run232 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties224 = new RunProperties();
            RunFonts runFonts245 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize272 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript125 = new FontSizeComplexScript() { Val = "16" };

            runProperties224.Append(runFonts245);
            runProperties224.Append(fontSize272);
            runProperties224.Append(fontSizeComplexScript125);
            Text text214 = new Text();
            text214.Text = "「";

            run232.Append(runProperties224);
            run232.Append(text214);

            Run run233 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties225 = new RunProperties();
            FontSize fontSize273 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript126 = new FontSizeComplexScript() { Val = "16" };

            runProperties225.Append(fontSize273);
            runProperties225.Append(fontSizeComplexScript126);
            Text text215 = new Text();
            text215.Text = "BA4002";

            run233.Append(runProperties225);
            run233.Append(text215);

            Run run234 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties226 = new RunProperties();
            RunFonts runFonts246 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize274 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript127 = new FontSizeComplexScript() { Val = "16" };

            runProperties226.Append(runFonts246);
            runProperties226.Append(fontSize274);
            runProperties226.Append(fontSizeComplexScript127);
            Text text216 = new Text();
            text216.Text = "內部稽核報告總結";

            run234.Append(runProperties226);
            run234.Append(text216);

            Run run235 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties227 = new RunProperties();
            RunFonts runFonts247 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize275 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript128 = new FontSizeComplexScript() { Val = "16" };

            runProperties227.Append(runFonts247);
            runProperties227.Append(fontSize275);
            runProperties227.Append(fontSizeComplexScript128);
            Text text217 = new Text();
            text217.Text = "_";

            run235.Append(runProperties227);
            run235.Append(text217);

            Run run236 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties228 = new RunProperties();
            RunFonts runFonts248 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize276 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript129 = new FontSizeComplexScript() { Val = "16" };

            runProperties228.Append(runFonts248);
            runProperties228.Append(fontSize276);
            runProperties228.Append(fontSizeComplexScript129);
            Text text218 = new Text();
            text218.Text = "肆、二、覆查前次主要檢查缺失」乙節規範若";

            run236.Append(runProperties228);
            run236.Append(text218);

            Run run237 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties229 = new RunProperties();
            RunFonts runFonts249 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize277 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript130 = new FontSizeComplexScript() { Val = "16" };

            runProperties229.Append(runFonts249);
            runProperties229.Append(fontSize277);
            runProperties229.Append(fontSizeComplexScript130);
            Text text219 = new Text();
            text219.Text = "前次無主要檢查缺失";

            run237.Append(runProperties229);
            run237.Append(text219);

            Run run238 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties230 = new RunProperties();
            RunFonts runFonts250 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize278 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript131 = new FontSizeComplexScript() { Val = "16" };

            runProperties230.Append(runFonts250);
            runProperties230.Append(fontSize278);
            runProperties230.Append(fontSizeComplexScript131);
            Text text220 = new Text();
            text220.Text = "可無須揭露";

            run238.Append(runProperties230);
            run238.Append(text220);

            Run run239 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties231 = new RunProperties();
            RunFonts runFonts251 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize279 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript132 = new FontSizeComplexScript() { Val = "16" };

            runProperties231.Append(runFonts251);
            runProperties231.Append(fontSize279);
            runProperties231.Append(fontSizeComplexScript132);
            Text text221 = new Text();
            text221.Text = "，";

            run239.Append(runProperties231);
            run239.Append(text221);
            ProofError proofError21 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run240 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties232 = new RunProperties();
            RunFonts runFonts252 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize280 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript133 = new FontSizeComplexScript() { Val = "16" };

            runProperties232.Append(runFonts252);
            runProperties232.Append(fontSize280);
            runProperties232.Append(fontSizeComplexScript133);
            Text text222 = new Text();
            text222.Text = "惟";

            run240.Append(runProperties232);
            run240.Append(text222);

            Run run241 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties233 = new RunProperties();
            RunFonts runFonts253 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize281 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript134 = new FontSizeComplexScript() { Val = "16" };

            runProperties233.Append(runFonts253);
            runProperties233.Append(fontSize281);
            runProperties233.Append(fontSizeComplexScript134);
            Text text223 = new Text();
            text223.Text = "請查核";

            run241.Append(runProperties233);
            run241.Append(text223);
            ProofError proofError22 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run242 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties234 = new RunProperties();
            RunFonts runFonts254 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize282 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript135 = new FontSizeComplexScript() { Val = "16" };

            runProperties234.Append(runFonts254);
            runProperties234.Append(fontSize282);
            runProperties234.Append(fontSizeComplexScript135);
            Text text224 = new Text();
            text224.Text = "人員確實於";

            run242.Append(runProperties234);
            run242.Append(text224);

            Run run243 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties235 = new RunProperties();
            RunFonts runFonts255 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize283 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript136 = new FontSizeComplexScript() { Val = "16" };

            runProperties235.Append(runFonts255);
            runProperties235.Append(fontSize283);
            runProperties235.Append(fontSizeComplexScript136);
            Text text225 = new Text();
            text225.Text = "本段";

            run243.Append(runProperties235);
            run243.Append(text225);

            Run run244 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties236 = new RunProperties();
            RunFonts runFonts256 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize284 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript137 = new FontSizeComplexScript() { Val = "16" };

            runProperties236.Append(runFonts256);
            runProperties236.Append(fontSize284);
            runProperties236.Append(fontSizeComplexScript137);
            Text text226 = new Text();
            text226.Text = "表述前次主要檢查缺失為";

            run244.Append(runProperties236);
            run244.Append(text226);

            Run run245 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties237 = new RunProperties();
            RunFonts runFonts257 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize285 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript138 = new FontSizeComplexScript() { Val = "16" };

            runProperties237.Append(runFonts257);
            runProperties237.Append(fontSize285);
            runProperties237.Append(fontSizeComplexScript138);
            Text text227 = new Text();
            text227.Text = "\"";

            run245.Append(runProperties237);
            run245.Append(text227);

            Run run246 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties238 = new RunProperties();
            RunFonts runFonts258 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize286 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript139 = new FontSizeComplexScript() { Val = "16" };

            runProperties238.Append(runFonts258);
            runProperties238.Append(fontSize286);
            runProperties238.Append(fontSizeComplexScript139);
            Text text228 = new Text();
            text228.Text = "無";

            run246.Append(runProperties238);
            run246.Append(text228);

            Run run247 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties239 = new RunProperties();
            RunFonts runFonts259 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize287 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript140 = new FontSizeComplexScript() { Val = "16" };

            runProperties239.Append(runFonts259);
            runProperties239.Append(fontSize287);
            runProperties239.Append(fontSizeComplexScript140);
            Text text229 = new Text();
            text229.Text = "\"";

            run247.Append(runProperties239);
            run247.Append(text229);

            Run run248 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "00051944" };

            RunProperties runProperties240 = new RunProperties();
            RunFonts runFonts260 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize288 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript141 = new FontSizeComplexScript() { Val = "16" };

            runProperties240.Append(runFonts260);
            runProperties240.Append(fontSize288);
            runProperties240.Append(fontSizeComplexScript141);
            Text text230 = new Text();
            text230.Text = "，以作為檢視完成之紀錄";

            run248.Append(runProperties240);
            run248.Append(text230);

            Run run249 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties241 = new RunProperties();
            RunFonts runFonts261 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize289 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript142 = new FontSizeComplexScript() { Val = "16" };

            runProperties241.Append(runFonts261);
            runProperties241.Append(fontSize289);
            runProperties241.Append(fontSizeComplexScript142);
            Text text231 = new Text();
            text231.Text = "。";

            run249.Append(runProperties241);
            run249.Append(text231);

            paragraph89.Append(paragraphProperties82);
            paragraph89.Append(run231);
            paragraph89.Append(run232);
            paragraph89.Append(run233);
            paragraph89.Append(run234);
            paragraph89.Append(run235);
            paragraph89.Append(run236);
            paragraph89.Append(run237);
            paragraph89.Append(run238);
            paragraph89.Append(run239);
            paragraph89.Append(proofError21);
            paragraph89.Append(run240);
            paragraph89.Append(run241);
            paragraph89.Append(proofError22);
            paragraph89.Append(run242);
            paragraph89.Append(run243);
            paragraph89.Append(run244);
            paragraph89.Append(run245);
            paragraph89.Append(run246);
            paragraph89.Append(run247);
            paragraph89.Append(run248);
            paragraph89.Append(run249);

            footnote4.Append(paragraph89);

            Footnote footnote5 = new Footnote() { Id = 3 };

            Paragraph paragraph90 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F03C81", RsidParagraphProperties = "00F03C81", RsidRunAdditionDefault = "00F03C81", ParagraphId = "2CF712F1", TextId = "77777777" };

            ParagraphProperties paragraphProperties83 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines36 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation166 = new Indentation() { Start = "158", Hanging = "158", HangingChars = 99 };

            ParagraphMarkRunProperties paragraphMarkRunProperties76 = new ParagraphMarkRunProperties();
            FontSize fontSize290 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript143 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties76.Append(fontSize290);
            paragraphMarkRunProperties76.Append(fontSizeComplexScript143);

            paragraphProperties83.Append(spacingBetweenLines36);
            paragraphProperties83.Append(indentation166);
            paragraphProperties83.Append(paragraphMarkRunProperties76);

            Run run250 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties242 = new RunProperties();
            RunStyle runStyle13 = new RunStyle() { Val = "a9" };
            FontSize fontSize291 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript144 = new FontSizeComplexScript() { Val = "16" };

            runProperties242.Append(runStyle13);
            runProperties242.Append(fontSize291);
            runProperties242.Append(fontSizeComplexScript144);
            FootnoteReferenceMark footnoteReferenceMark3 = new FootnoteReferenceMark();

            run250.Append(runProperties242);
            run250.Append(footnoteReferenceMark3);

            Run run251 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties243 = new RunProperties();
            FontSize fontSize292 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript145 = new FontSizeComplexScript() { Val = "16" };

            runProperties243.Append(fontSize292);
            runProperties243.Append(fontSizeComplexScript145);
            Text text232 = new Text();
            text232.Text = "a)";

            run251.Append(runProperties243);
            run251.Append(text232);

            Run run252 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties244 = new RunProperties();
            FontSize fontSize293 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript146 = new FontSizeComplexScript() { Val = "16" };

            runProperties244.Append(fontSize293);
            runProperties244.Append(fontSizeComplexScript146);
            Text text233 = new Text();
            text233.Text = "業務申辦或缺失改善經主管機關核示事項之遵循情形：係指確認受檢單位於業務申請或缺失改善時，其向主管機關申報之承諾事項，以及主管機關來函之核示內容之遵循情形。";

            run252.Append(runProperties244);
            run252.Append(text233);

            paragraph90.Append(paragraphProperties83);
            paragraph90.Append(run250);
            paragraph90.Append(run251);
            paragraph90.Append(run252);

            Paragraph paragraph91 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F03C81", RsidParagraphProperties = "00F03C81", RsidRunAdditionDefault = "00F03C81", ParagraphId = "1D9E74CE", TextId = "77777777" };

            ParagraphProperties paragraphProperties84 = new ParagraphProperties();
            SpacingBetweenLines spacingBetweenLines37 = new SpacingBetweenLines() { Line = "240", LineRule = LineSpacingRuleValues.Exact };
            Indentation indentation167 = new Indentation() { Start = "195", StartCharacters = 29, Hanging = "125", HangingChars = 78 };

            ParagraphMarkRunProperties paragraphMarkRunProperties77 = new ParagraphMarkRunProperties();
            FontSize fontSize294 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript147 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties77.Append(fontSize294);
            paragraphMarkRunProperties77.Append(fontSizeComplexScript147);

            paragraphProperties84.Append(spacingBetweenLines37);
            paragraphProperties84.Append(indentation167);
            paragraphProperties84.Append(paragraphMarkRunProperties77);

            Run run253 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties245 = new RunProperties();
            FontSize fontSize295 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript148 = new FontSizeComplexScript() { Val = "16" };

            runProperties245.Append(fontSize295);
            runProperties245.Append(fontSizeComplexScript148);
            Text text234 = new Text();
            text234.Text = "b)";

            run253.Append(runProperties245);
            run253.Append(text234);

            Run run254 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties246 = new RunProperties();
            FontSize fontSize296 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript149 = new FontSizeComplexScript() { Val = "16" };

            runProperties246.Append(fontSize296);
            runProperties246.Append(fontSizeComplexScript149);
            Text text235 = new Text();
            text235.Text = "結構型商品異常客訴案件：依據";

            run254.Append(runProperties246);
            run254.Append(text235);

            Run run255 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties247 = new RunProperties();
            FontSize fontSize297 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript150 = new FontSizeComplexScript() { Val = "16" };

            runProperties247.Append(fontSize297);
            runProperties247.Append(fontSizeComplexScript150);
            Text text236 = new Text();
            text236.Text = "106.6.5";

            run255.Append(runProperties247);
            run255.Append(text236);

            Run run256 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties248 = new RunProperties();
            FontSize fontSize298 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript151 = new FontSizeComplexScript() { Val = "16" };

            runProperties248.Append(fontSize298);
            runProperties248.Append(fontSizeComplexScript151);
            Text text237 = new Text();
            text237.Text = "檢局（控）字第";

            run256.Append(runProperties248);
            run256.Append(text237);

            Run run257 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties249 = new RunProperties();
            FontSize fontSize299 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript152 = new FontSizeComplexScript() { Val = "16" };

            runProperties249.Append(fontSize299);
            runProperties249.Append(fontSizeComplexScript152);
            Text text238 = new Text();
            text238.Text = "1060152199";

            run257.Append(runProperties249);
            run257.Append(text238);

            Run run258 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties250 = new RunProperties();
            FontSize fontSize300 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript153 = new FontSizeComplexScript() { Val = "16" };

            runProperties250.Append(fontSize300);
            runProperties250.Append(fontSizeComplexScript153);
            Text text239 = new Text();
            text239.Text = "號函要求，鑒於近來銀行因銷售結構型商品衍生之";

            run258.Append(runProperties250);
            run258.Append(text239);
            ProofError proofError23 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run259 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties251 = new RunProperties();
            FontSize fontSize301 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript154 = new FontSizeComplexScript() { Val = "16" };

            runProperties251.Append(fontSize301);
            runProperties251.Append(fontSizeComplexScript154);
            Text text240 = new Text();
            text240.Text = "客訴日增";

            run259.Append(runProperties251);
            run259.Append(text240);
            ProofError proofError24 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run260 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties252 = new RunProperties();
            FontSize fontSize302 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript155 = new FontSizeComplexScript() { Val = "16" };

            runProperties252.Append(fontSize302);
            runProperties252.Append(fontSizeComplexScript155);
            Text text241 = new Text();
            text241.Text = "，請稽核單位加強抽查異常客訴案件，內部稽核查核客訴案件辦理情形，將列為檢查局對銀行稽核工作考核要項。";

            run260.Append(runProperties252);
            run260.Append(text241);

            paragraph91.Append(paragraphProperties84);
            paragraph91.Append(run253);
            paragraph91.Append(run254);
            paragraph91.Append(run255);
            paragraph91.Append(run256);
            paragraph91.Append(run257);
            paragraph91.Append(run258);
            paragraph91.Append(proofError23);
            paragraph91.Append(run259);
            paragraph91.Append(proofError24);
            paragraph91.Append(run260);

            Paragraph paragraph92 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F03C81", RsidParagraphProperties = "00F03C81", RsidRunAdditionDefault = "00F03C81", ParagraphId = "62F3C0BA", TextId = "77777777" };

            ParagraphProperties paragraphProperties85 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId24 = new ParagraphStyleId() { Val = "a7" };
            Indentation indentation168 = new Indentation() { FirstLine = "80", FirstLineChars = 50 };

            paragraphProperties85.Append(paragraphStyleId24);
            paragraphProperties85.Append(indentation168);

            Run run261 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties253 = new RunProperties();
            RunFonts runFonts262 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };
            FontSize fontSize303 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript156 = new FontSizeComplexScript() { Val = "16" };

            runProperties253.Append(runFonts262);
            runProperties253.Append(fontSize303);
            runProperties253.Append(fontSizeComplexScript156);
            Text text242 = new Text();
            text242.Text = "c)";

            run261.Append(runProperties253);
            run261.Append(text242);

            Run run262 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties254 = new RunProperties();
            FontSize fontSize304 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript157 = new FontSizeComplexScript() { Val = "16" };

            runProperties254.Append(fontSize304);
            runProperties254.Append(fontSizeComplexScript157);
            Text text243 = new Text();
            text243.Text = "部分重大內控缺失係透過";

            run262.Append(runProperties254);
            run262.Append(text243);
            ProofError proofError25 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run263 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties255 = new RunProperties();
            FontSize fontSize305 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript158 = new FontSizeComplexScript() { Val = "16" };

            runProperties255.Append(fontSize305);
            runProperties255.Append(fontSizeComplexScript158);
            Text text244 = new Text();
            text244.Text = "客訴或檢舉";

            run263.Append(runProperties255);
            run263.Append(text244);
            ProofError proofError26 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run264 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties256 = new RunProperties();
            FontSize fontSize306 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript159 = new FontSizeComplexScript() { Val = "16" };

            runProperties256.Append(fontSize306);
            runProperties256.Append(fontSizeComplexScript159);
            Text text245 = new Text();
            text245.Text = "事件發現，建議亦可評估是否納為查核重點。";

            run264.Append(runProperties256);
            run264.Append(text245);

            paragraph92.Append(paragraphProperties85);
            paragraph92.Append(run261);
            paragraph92.Append(run262);
            paragraph92.Append(proofError25);
            paragraph92.Append(run263);
            paragraph92.Append(proofError26);
            paragraph92.Append(run264);

            footnote5.Append(paragraph90);
            footnote5.Append(paragraph91);
            footnote5.Append(paragraph92);

            Footnote footnote6 = new Footnote() { Id = 4 };

            Paragraph paragraph93 = new Paragraph() { RsidParagraphMarkRevision = "00D62F73", RsidParagraphAddition = "00F87EAC", RsidRunAdditionDefault = "00F87EAC", ParagraphId = "3C357AC2", TextId = "77777777" };

            ParagraphProperties paragraphProperties86 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId25 = new ParagraphStyleId() { Val = "a7" };

            ParagraphMarkRunProperties paragraphMarkRunProperties78 = new ParagraphMarkRunProperties();
            RunFonts runFonts263 = new RunFonts() { EastAsia = "新細明體" };
            Kern kern8 = new Kern() { Val = (UInt32Value)0U };
            FontSize fontSize307 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript160 = new FontSizeComplexScript() { Val = "16" };
            Underline underline13 = new Underline() { Val = UnderlineValues.Single };

            paragraphMarkRunProperties78.Append(runFonts263);
            paragraphMarkRunProperties78.Append(kern8);
            paragraphMarkRunProperties78.Append(fontSize307);
            paragraphMarkRunProperties78.Append(fontSizeComplexScript160);
            paragraphMarkRunProperties78.Append(underline13);

            paragraphProperties86.Append(paragraphStyleId25);
            paragraphProperties86.Append(paragraphMarkRunProperties78);

            Run run265 = new Run() { RsidRunProperties = "00D62F73" };

            RunProperties runProperties257 = new RunProperties();
            RunStyle runStyle14 = new RunStyle() { Val = "a9" };
            FontSize fontSize308 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript161 = new FontSizeComplexScript() { Val = "16" };

            runProperties257.Append(runStyle14);
            runProperties257.Append(fontSize308);
            runProperties257.Append(fontSizeComplexScript161);
            FootnoteReferenceMark footnoteReferenceMark4 = new FootnoteReferenceMark();

            run265.Append(runProperties257);
            run265.Append(footnoteReferenceMark4);

            Run run266 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties258 = new RunProperties();
            FontSize fontSize309 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript162 = new FontSizeComplexScript() { Val = "16" };

            runProperties258.Append(fontSize309);
            runProperties258.Append(fontSizeComplexScript162);
            Text text246 = new Text();
            text246.Text = "[";

            run266.Append(runProperties258);
            run266.Append(text246);

            Run run267 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties259 = new RunProperties();
            RunFonts runFonts264 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize310 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript163 = new FontSizeComplexScript() { Val = "16" };

            runProperties259.Append(runFonts264);
            runProperties259.Append(fontSize310);
            runProperties259.Append(fontSizeComplexScript163);
            Text text247 = new Text();
            text247.Text = "國內分行";

            run267.Append(runProperties259);
            run267.Append(text247);

            Run run268 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties260 = new RunProperties();
            FontSize fontSize311 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript164 = new FontSizeComplexScript() { Val = "16" };

            runProperties260.Append(fontSize311);
            runProperties260.Append(fontSizeComplexScript164);
            Text text248 = new Text();
            text248.Text = "]";

            run268.Append(runProperties260);
            run268.Append(text248);

            Run run269 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties261 = new RunProperties();
            RunFonts runFonts265 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize312 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript165 = new FontSizeComplexScript() { Val = "16" };

            runProperties261.Append(runFonts265);
            runProperties261.Append(fontSize312);
            runProperties261.Append(fontSizeComplexScript165);
            Text text249 = new Text();
            text249.Text = "及";

            run269.Append(runProperties261);
            run269.Append(text249);

            Run run270 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties262 = new RunProperties();
            FontSize fontSize313 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript166 = new FontSizeComplexScript() { Val = "16" };

            runProperties262.Append(fontSize313);
            runProperties262.Append(fontSizeComplexScript166);
            Text text250 = new Text();
            text250.Text = "[";

            run270.Append(runProperties262);
            run270.Append(text250);
            ProofError proofError27 = new ProofError() { Type = ProofingErrorValues.GrammarStart };

            Run run271 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties263 = new RunProperties();
            RunFonts runFonts266 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize314 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript167 = new FontSizeComplexScript() { Val = "16" };

            runProperties263.Append(runFonts266);
            runProperties263.Append(fontSize314);
            runProperties263.Append(fontSizeComplexScript167);
            Text text251 = new Text();
            text251.Text = "個";

            run271.Append(runProperties263);
            run271.Append(text251);
            ProofError proofError28 = new ProofError() { Type = ProofingErrorValues.GrammarEnd };

            Run run272 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties264 = new RunProperties();
            RunFonts runFonts267 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize315 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript168 = new FontSizeComplexScript() { Val = "16" };

            runProperties264.Append(runFonts267);
            runProperties264.Append(fontSize315);
            runProperties264.Append(fontSizeComplexScript168);
            Text text252 = new Text();
            text252.Text = "金區域中心";

            run272.Append(runProperties264);
            run272.Append(text252);

            Run run273 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties265 = new RunProperties();
            FontSize fontSize316 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript169 = new FontSizeComplexScript() { Val = "16" };

            runProperties265.Append(fontSize316);
            runProperties265.Append(fontSizeComplexScript169);
            Text text253 = new Text();
            text253.Text = "]";

            run273.Append(runProperties265);
            run273.Append(text253);

            Run run274 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties266 = new RunProperties();
            RunFonts runFonts268 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize317 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript170 = new FontSizeComplexScript() { Val = "16" };

            runProperties266.Append(runFonts268);
            runProperties266.Append(fontSize317);
            runProperties266.Append(fontSizeComplexScript170);
            Text text254 = new Text();
            text254.Text = "最近一次年度內部稽核風險評估結果為高風險者，應";

            run274.Append(runProperties266);
            run274.Append(text254);

            Run run275 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties267 = new RunProperties();
            RunFonts runFonts269 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize318 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript171 = new FontSizeComplexScript() { Val = "16" };
            Languages languages23 = new Languages() { EastAsia = "zh-HK" };

            runProperties267.Append(runFonts269);
            runProperties267.Append(fontSize318);
            runProperties267.Append(fontSizeComplexScript171);
            runProperties267.Append(languages23);
            Text text255 = new Text();
            text255.Text = "於本項說明";

            run275.Append(runProperties267);
            run275.Append(text255);

            Run run276 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties268 = new RunProperties();
            RunFonts runFonts270 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize319 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript172 = new FontSizeComplexScript() { Val = "16" };

            runProperties268.Append(runFonts270);
            runProperties268.Append(fontSize319);
            runProperties268.Append(fontSizeComplexScript172);
            Text text256 = new Text();
            text256.Text = "：";

            run276.Append(runProperties268);
            run276.Append(text256);

            Run run277 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties269 = new RunProperties();
            RunFonts runFonts271 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize320 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript173 = new FontSizeComplexScript() { Val = "16" };
            Languages languages24 = new Languages() { EastAsia = "zh-HK" };

            runProperties269.Append(runFonts271);
            runProperties269.Append(fontSize320);
            runProperties269.Append(fontSizeComplexScript173);
            runProperties269.Append(languages24);
            Text text257 = new Text();
            text257.Text = "該";

            run277.Append(runProperties269);
            run277.Append(text257);

            Run run278 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties270 = new RunProperties();
            FontSize fontSize321 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript174 = new FontSizeComplexScript() { Val = "16" };

            runProperties270.Append(fontSize321);
            runProperties270.Append(fontSizeComplexScript174);
            Text text258 = new Text();
            text258.Text = "[";

            run278.Append(runProperties270);
            run278.Append(text258);

            Run run279 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties271 = new RunProperties();
            RunFonts runFonts272 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize322 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript175 = new FontSizeComplexScript() { Val = "16" };

            runProperties271.Append(runFonts272);
            runProperties271.Append(fontSize322);
            runProperties271.Append(fontSizeComplexScript175);
            Text text259 = new Text();
            text259.Text = "國內分行";

            run279.Append(runProperties271);
            run279.Append(text259);

            Run run280 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties272 = new RunProperties();
            FontSize fontSize323 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript176 = new FontSizeComplexScript() { Val = "16" };

            runProperties272.Append(fontSize323);
            runProperties272.Append(fontSizeComplexScript176);
            Text text260 = new Text();
            text260.Text = "]";

            run280.Append(runProperties272);
            run280.Append(text260);

            Run run281 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties273 = new RunProperties();
            RunFonts runFonts273 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize324 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript177 = new FontSizeComplexScript() { Val = "16" };
            Languages languages25 = new Languages() { EastAsia = "zh-HK" };

            runProperties273.Append(runFonts273);
            runProperties273.Append(fontSize324);
            runProperties273.Append(fontSizeComplexScript177);
            runProperties273.Append(languages25);
            Text text261 = new Text();
            text261.Text = "或";

            run281.Append(runProperties273);
            run281.Append(text261);

            Run run282 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties274 = new RunProperties();
            FontSize fontSize325 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript178 = new FontSizeComplexScript() { Val = "16" };

            runProperties274.Append(fontSize325);
            runProperties274.Append(fontSizeComplexScript178);
            Text text262 = new Text();
            text262.Text = "[";

            run282.Append(runProperties274);
            run282.Append(text262);

            Run run283 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties275 = new RunProperties();
            RunFonts runFonts274 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize326 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript179 = new FontSizeComplexScript() { Val = "16" };

            runProperties275.Append(runFonts274);
            runProperties275.Append(fontSize326);
            runProperties275.Append(fontSizeComplexScript179);
            Text text263 = new Text();
            text263.Text = "個金區域中心";

            run283.Append(runProperties275);
            run283.Append(text263);

            Run run284 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties276 = new RunProperties();
            FontSize fontSize327 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript180 = new FontSizeComplexScript() { Val = "16" };

            runProperties276.Append(fontSize327);
            runProperties276.Append(fontSizeComplexScript180);
            Text text264 = new Text();
            text264.Text = "]";

            run284.Append(runProperties276);
            run284.Append(text264);

            Run run285 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties277 = new RunProperties();
            RunFonts runFonts275 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize328 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript181 = new FontSizeComplexScript() { Val = "16" };
            Languages languages26 = new Languages() { EastAsia = "zh-HK" };

            runProperties277.Append(runFonts275);
            runProperties277.Append(fontSize328);
            runProperties277.Append(fontSizeComplexScript181);
            runProperties277.Append(languages26);
            Text text265 = new Text();
            text265.Text = "最近一次年度";

            run285.Append(runProperties277);
            run285.Append(text265);

            Run run286 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties278 = new RunProperties();
            RunFonts runFonts276 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize329 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript182 = new FontSizeComplexScript() { Val = "16" };

            runProperties278.Append(runFonts276);
            runProperties278.Append(fontSize329);
            runProperties278.Append(fontSizeComplexScript182);
            Text text266 = new Text();
            text266.Text = "風險評估結果";

            run286.Append(runProperties278);
            run286.Append(text266);

            Run run287 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties279 = new RunProperties();
            RunFonts runFonts277 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize330 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript183 = new FontSizeComplexScript() { Val = "16" };
            Languages languages27 = new Languages() { EastAsia = "zh-HK" };

            runProperties279.Append(runFonts277);
            runProperties279.Append(fontSize330);
            runProperties279.Append(fontSizeComplexScript183);
            runProperties279.Append(languages27);
            Text text267 = new Text();
            text267.Text = "為高";

            run287.Append(runProperties279);
            run287.Append(text267);

            Run run288 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties280 = new RunProperties();
            RunFonts runFonts278 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize331 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript184 = new FontSizeComplexScript() { Val = "16" };

            runProperties280.Append(runFonts278);
            runProperties280.Append(fontSize331);
            runProperties280.Append(fontSizeComplexScript184);
            Text text268 = new Text();
            text268.Text = "，";

            run288.Append(runProperties280);
            run288.Append(text268);

            Run run289 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties281 = new RunProperties();
            RunFonts runFonts279 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize332 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript185 = new FontSizeComplexScript() { Val = "16" };
            Languages languages28 = new Languages() { EastAsia = "zh-HK" };

            runProperties281.Append(runFonts279);
            runProperties281.Append(fontSize332);
            runProperties281.Append(fontSizeComplexScript185);
            runProperties281.Append(languages28);
            Text text269 = new Text();
            text269.Text = "本次一般查核將辦理";

            run289.Append(runProperties281);
            run289.Append(text269);

            Run run290 = new Run() { RsidRunProperties = "00D62F73", RsidRunAddition = "009A6EFB" };

            RunProperties runProperties282 = new RunProperties();
            RunFonts runFonts280 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "標楷體", HighAnsi = "標楷體" };
            FontSize fontSize333 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript186 = new FontSizeComplexScript() { Val = "16" };

            runProperties282.Append(runFonts280);
            runProperties282.Append(fontSize333);
            runProperties282.Append(fontSizeComplexScript186);
            Text text270 = new Text();
            text270.Text = "單位管理階層風險意識評估。";

            run290.Append(runProperties282);
            run290.Append(text270);

            paragraph93.Append(paragraphProperties86);
            paragraph93.Append(run265);
            paragraph93.Append(run266);
            paragraph93.Append(run267);
            paragraph93.Append(run268);
            paragraph93.Append(run269);
            paragraph93.Append(run270);
            paragraph93.Append(proofError27);
            paragraph93.Append(run271);
            paragraph93.Append(proofError28);
            paragraph93.Append(run272);
            paragraph93.Append(run273);
            paragraph93.Append(run274);
            paragraph93.Append(run275);
            paragraph93.Append(run276);
            paragraph93.Append(run277);
            paragraph93.Append(run278);
            paragraph93.Append(run279);
            paragraph93.Append(run280);
            paragraph93.Append(run281);
            paragraph93.Append(run282);
            paragraph93.Append(run283);
            paragraph93.Append(run284);
            paragraph93.Append(run285);
            paragraph93.Append(run286);
            paragraph93.Append(run287);
            paragraph93.Append(run288);
            paragraph93.Append(run289);
            paragraph93.Append(run290);

            footnote6.Append(paragraph93);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);
            footnotes1.Append(footnote3);
            footnotes1.Append(footnote4);
            footnotes1.Append(footnote5);
            footnotes1.Append(footnote6);

            footnotesPart1.Footnotes = footnotes1;
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
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex3);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex4);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent1Color1.Append(rgbColorModelHex5);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex6);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex7);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex8);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent5Color1.Append(rgbColorModelHex9);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex10);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex11);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex12);

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
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

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
            majorFont1.Append(supplementalFont31);
            majorFont1.Append(supplementalFont32);
            majorFont1.Append(supplementalFont33);
            majorFont1.Append(supplementalFont34);
            majorFont1.Append(supplementalFont35);
            majorFont1.Append(supplementalFont36);
            majorFont1.Append(supplementalFont37);
            majorFont1.Append(supplementalFont38);
            majorFont1.Append(supplementalFont39);
            majorFont1.Append(supplementalFont40);
            majorFont1.Append(supplementalFont41);
            majorFont1.Append(supplementalFont42);
            majorFont1.Append(supplementalFont43);
            majorFont1.Append(supplementalFont44);
            majorFont1.Append(supplementalFont45);
            majorFont1.Append(supplementalFont46);
            majorFont1.Append(supplementalFont47);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Jpan", Typeface = "游明朝" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Hans", Typeface = "等线" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont61 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont62 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont63 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont64 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont65 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont66 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont67 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont68 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont69 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont70 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont71 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont72 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont73 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont74 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont75 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont76 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont77 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };
            A.SupplementalFont supplementalFont78 = new A.SupplementalFont() { Script = "Armn", Typeface = "Arial" };
            A.SupplementalFont supplementalFont79 = new A.SupplementalFont() { Script = "Bugi", Typeface = "Leelawadee UI" };
            A.SupplementalFont supplementalFont80 = new A.SupplementalFont() { Script = "Bopo", Typeface = "Microsoft JhengHei" };
            A.SupplementalFont supplementalFont81 = new A.SupplementalFont() { Script = "Java", Typeface = "Javanese Text" };
            A.SupplementalFont supplementalFont82 = new A.SupplementalFont() { Script = "Lisu", Typeface = "Segoe UI" };
            A.SupplementalFont supplementalFont83 = new A.SupplementalFont() { Script = "Mymr", Typeface = "Myanmar Text" };
            A.SupplementalFont supplementalFont84 = new A.SupplementalFont() { Script = "Nkoo", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont85 = new A.SupplementalFont() { Script = "Olck", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont86 = new A.SupplementalFont() { Script = "Osma", Typeface = "Ebrima" };
            A.SupplementalFont supplementalFont87 = new A.SupplementalFont() { Script = "Phag", Typeface = "Phagspa" };
            A.SupplementalFont supplementalFont88 = new A.SupplementalFont() { Script = "Syrn", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont89 = new A.SupplementalFont() { Script = "Syrj", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont90 = new A.SupplementalFont() { Script = "Syre", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont91 = new A.SupplementalFont() { Script = "Sora", Typeface = "Nirmala UI" };
            A.SupplementalFont supplementalFont92 = new A.SupplementalFont() { Script = "Tale", Typeface = "Microsoft Tai Le" };
            A.SupplementalFont supplementalFont93 = new A.SupplementalFont() { Script = "Talu", Typeface = "Microsoft New Tai Lue" };
            A.SupplementalFont supplementalFont94 = new A.SupplementalFont() { Script = "Tfng", Typeface = "Ebrima" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
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
            minorFont1.Append(supplementalFont61);
            minorFont1.Append(supplementalFont62);
            minorFont1.Append(supplementalFont63);
            minorFont1.Append(supplementalFont64);
            minorFont1.Append(supplementalFont65);
            minorFont1.Append(supplementalFont66);
            minorFont1.Append(supplementalFont67);
            minorFont1.Append(supplementalFont68);
            minorFont1.Append(supplementalFont69);
            minorFont1.Append(supplementalFont70);
            minorFont1.Append(supplementalFont71);
            minorFont1.Append(supplementalFont72);
            minorFont1.Append(supplementalFont73);
            minorFont1.Append(supplementalFont74);
            minorFont1.Append(supplementalFont75);
            minorFont1.Append(supplementalFont76);
            minorFont1.Append(supplementalFont77);
            minorFont1.Append(supplementalFont78);
            minorFont1.Append(supplementalFont79);
            minorFont1.Append(supplementalFont80);
            minorFont1.Append(supplementalFont81);
            minorFont1.Append(supplementalFont82);
            minorFont1.Append(supplementalFont83);
            minorFont1.Append(supplementalFont84);
            minorFont1.Append(supplementalFont85);
            minorFont1.Append(supplementalFont86);
            minorFont1.Append(supplementalFont87);
            minorFont1.Append(supplementalFont88);
            minorFont1.Append(supplementalFont89);
            minorFont1.Append(supplementalFont90);
            minorFont1.Append(supplementalFont91);
            minorFont1.Append(supplementalFont92);
            minorFont1.Append(supplementalFont93);
            minorFont1.Append(supplementalFont94);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor1);

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

            fillStyleList1.Append(solidFill3);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline3 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor8);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash2);
            outline3.Append(miter2);

            A.Outline outline4 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor9);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline4.Append(solidFill5);
            outline4.Append(presetDash3);
            outline4.Append(miter3);

            A.Outline outline5 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill6.Append(schemeColor10);
            A.PresetDash presetDash4 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter4 = new A.Miter() { Limit = 800000 };

            outline5.Append(solidFill6);
            outline5.Append(presetDash4);
            outline5.Append(miter4);

            lineStyleList1.Append(outline3);
            lineStyleList1.Append(outline4);
            lineStyleList1.Append(outline5);

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

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex13.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill7 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill7.Append(schemeColor11);

            A.SolidFill solidFill8 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill8.Append(schemeColor12);

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

            backgroundFillStyleList1.Append(solidFill7);
            backgroundFillStyleList1.Append(solidFill8);
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

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Z00002110";
            document.PackageProperties.Title = "內部稽核查核計劃(範本)";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Revision = "4";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2023-08-11T05:53:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2023-08-11T08:09:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "余亭妍";
            document.PackageProperties.LastPrinted = System.Xml.XmlConvert.ToDateTime("2023-03-16T09:17:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
        }


    }
}
