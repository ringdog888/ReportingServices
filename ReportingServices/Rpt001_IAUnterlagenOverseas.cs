using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using M = DocumentFormat.OpenXml.Math;
using A = DocumentFormat.OpenXml.Drawing;
using System.Data;
using System.Collections.Generic;

namespace Rpt001
{
    public class GeneratedClass
    {
        // Data Source
        public DataTable dt_L1 { get; set; }
        public DataTable dt_L2 { get; set; }
        public string connectionString { get; set; }

        //Creates Report Tool
        ReportingServices.RptTool RptTool = new ReportingServices.RptTool();

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

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId8");
            GenerateThemePart1Content(themePart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId7");
            GenerateFontTablePart1Content(fontTablePart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            FooterPart footerPart1 = mainDocumentPart1.AddNewPart<FooterPart>("rId6");
            GenerateFooterPart1Content(footerPart1);

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
            totalTime1.Text = "0";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "26";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "150";
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
            charactersWithSpaces1.Text = "175";
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

            for (int i = 0; i < dt_L1.Rows.Count; i++)
            {
                document1.Append(AddBody1(dt_L1.Rows[i],i));
            }

            mainDocumentPart1.Document = document1;
        }

        Body AddBody1(DataRow Row_L1, int RowIndex)
        {
            Body body1 = new Body();

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableWidth tableWidth1 = new TableWidth() { Width = "19300", Type = TableWidthUnitValues.Dxa };
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "5000" };
            GridColumn gridColumn2 = new GridColumn() { Width = "9300" };
            GridColumn gridColumn3 = new GridColumn() { Width = "5000" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00591921", RsidTableRowAddition = "008473F6", RsidTableRowProperties = "00591921" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(shading1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize1 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);

            paragraphProperties1.Append(paragraphMarkRunProperties1);

            paragraph1.Append(paragraphProperties1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "9300", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "000D0440", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts2 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize2 = new FontSize() { Val = "32" };

            paragraphMarkRunProperties2.Append(runFonts2);
            paragraphMarkRunProperties2.Append(fontSize2);

            paragraphProperties2.Append(justification1);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run1 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts3 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize3 = new FontSize() { Val = "32" };

            runProperties1.Append(runFonts3);
            runProperties1.Append(fontSize3);
            Text text1 = new Text();
            text1.Text = "查核紀錄表－";

            run1.Append(runProperties1);
            run1.Append(text1);

            Run run2 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize4 = new FontSize() { Val = "32" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(fontSize4);
            Text text2 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text2.Text = " ";

            run2.Append(runProperties2);
            run2.Append(text2);

            Run run4 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties4 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize6 = new FontSize() { Val = "32" };

            runProperties4.Append(runFonts6);
            runProperties4.Append(fontSize6);
            Text text4 = new Text();
            text4.Text = Row_L1["L1Name"].ToString();

            run4.Append(runProperties4);
            run4.Append(text4);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run1);
            paragraph2.Append(run2);
            paragraph2.Append(run4);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts7 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize7 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties3.Append(runFonts7);
            paragraphMarkRunProperties3.Append(fontSize7);

            paragraphProperties3.Append(paragraphMarkRunProperties3);

            paragraph3.Append(paragraphProperties3);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            TableRow tableRow2 = new TableRow() { RsidTableRowMarkRevision = "00591921", RsidTableRowAddition = "008473F6", RsidTableRowProperties = "00591921" };

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties4 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
            RunFonts runFonts8 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize8 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties4.Append(runFonts8);
            paragraphMarkRunProperties4.Append(fontSize8);

            paragraphProperties4.Append(paragraphMarkRunProperties4);

            paragraph4.Append(paragraphProperties4);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "9300", Type = TableWidthUnitValues.Dxa };
            Shading shading5 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties5.Append(tableCellWidth5);
            tableCellProperties5.Append(shading5);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "000D0440", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties5 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties5 = new ParagraphMarkRunProperties();
            RunFonts runFonts9 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize9 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties5.Append(runFonts9);
            paragraphMarkRunProperties5.Append(fontSize9);

            paragraphProperties5.Append(paragraphMarkRunProperties5);

            paragraph5.Append(paragraphProperties5);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph5);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading6 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties6.Append(tableCellWidth6);
            tableCellProperties6.Append(shading6);

            Paragraph paragraph6 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties6 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties6 = new ParagraphMarkRunProperties();
            RunFonts runFonts10 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize10 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties6.Append(runFonts10);
            paragraphMarkRunProperties6.Append(fontSize10);

            paragraphProperties6.Append(paragraphMarkRunProperties6);

            Run run5 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties5 = new RunProperties();
            RunFonts runFonts11 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize11 = new FontSize() { Val = "28" };

            runProperties5.Append(runFonts11);
            runProperties5.Append(fontSize11);
            Text text5 = new Text();
            text5.Text = "受查單位：";

            run5.Append(runProperties5);
            run5.Append(text5);

            paragraph6.Append(paragraphProperties6);
            paragraph6.Append(run5);

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph6);

            tableRow2.Append(tableCell4);
            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);

            TableRow tableRow3 = new TableRow() { RsidTableRowMarkRevision = "00591921", RsidTableRowAddition = "008473F6", RsidTableRowProperties = "00591921" };

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading7 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties7.Append(tableCellWidth7);
            tableCellProperties7.Append(shading7);

            Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties7 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties7 = new ParagraphMarkRunProperties();
            RunFonts runFonts12 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize12 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties7.Append(runFonts12);
            paragraphMarkRunProperties7.Append(fontSize12);

            paragraphProperties7.Append(paragraphMarkRunProperties7);

            paragraph7.Append(paragraphProperties7);

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph7);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "9300", Type = TableWidthUnitValues.Dxa };
            Shading shading8 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties8.Append(tableCellWidth8);
            tableCellProperties8.Append(shading8);

            Paragraph paragraph8 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties8 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties8 = new ParagraphMarkRunProperties();
            RunFonts runFonts13 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize13 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties8.Append(runFonts13);
            paragraphMarkRunProperties8.Append(fontSize13);

            paragraphProperties8.Append(justification2);
            paragraphProperties8.Append(paragraphMarkRunProperties8);

            Run run6 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties6 = new RunProperties();
            RunFonts runFonts14 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize14 = new FontSize() { Val = "28" };

            runProperties6.Append(runFonts14);
            runProperties6.Append(fontSize14);
            Text text6 = new Text();
            text6.Text = "實際查核日：";

            run6.Append(runProperties6);
            run6.Append(text6);

            Run run7 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties7 = new RunProperties();
            RunFonts runFonts15 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize15 = new FontSize() { Val = "28" };

            runProperties7.Append(runFonts15);
            runProperties7.Append(fontSize15);
            Text text7 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text7.Text = "      ";

            run7.Append(runProperties7);
            run7.Append(text7);

            Run run8 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties8 = new RunProperties();
            RunFonts runFonts16 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize16 = new FontSize() { Val = "28" };

            runProperties8.Append(runFonts16);
            runProperties8.Append(fontSize16);
            Text text8 = new Text();
            text8.Text = "年";

            run8.Append(runProperties8);
            run8.Append(text8);

            Run run9 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties9 = new RunProperties();
            RunFonts runFonts17 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize17 = new FontSize() { Val = "28" };

            runProperties9.Append(runFonts17);
            runProperties9.Append(fontSize17);
            Text text9 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text9.Text = "      ";

            run9.Append(runProperties9);
            run9.Append(text9);

            Run run10 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties10 = new RunProperties();
            RunFonts runFonts18 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize18 = new FontSize() { Val = "28" };

            runProperties10.Append(runFonts18);
            runProperties10.Append(fontSize18);
            Text text10 = new Text();
            text10.Text = "月";

            run10.Append(runProperties10);
            run10.Append(text10);

            Run run11 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties11 = new RunProperties();
            RunFonts runFonts19 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize19 = new FontSize() { Val = "28" };

            runProperties11.Append(runFonts19);
            runProperties11.Append(fontSize19);
            Text text11 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text11.Text = "      ";

            run11.Append(runProperties11);
            run11.Append(text11);

            Run run12 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties12 = new RunProperties();
            RunFonts runFonts20 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize20 = new FontSize() { Val = "28" };

            runProperties12.Append(runFonts20);
            runProperties12.Append(fontSize20);
            Text text12 = new Text();
            text12.Text = "日";

            run12.Append(runProperties12);
            run12.Append(text12);

            paragraph8.Append(paragraphProperties8);
            paragraph8.Append(run6);
            paragraph8.Append(run7);
            paragraph8.Append(run8);
            paragraph8.Append(run9);
            paragraph8.Append(run10);
            paragraph8.Append(run11);
            paragraph8.Append(run12);

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph8);

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading9 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties9.Append(tableCellWidth9);
            tableCellProperties9.Append(shading9);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties9 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties9 = new ParagraphMarkRunProperties();
            RunFonts runFonts21 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize21 = new FontSize() { Val = "28" };

            paragraphMarkRunProperties9.Append(runFonts21);
            paragraphMarkRunProperties9.Append(fontSize21);

            paragraphProperties9.Append(paragraphMarkRunProperties9);

            Run run13 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties13 = new RunProperties();
            RunFonts runFonts22 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };
            FontSize fontSize22 = new FontSize() { Val = "28" };

            runProperties13.Append(runFonts22);
            runProperties13.Append(fontSize22);
            Text text13 = new Text();
            text13.Text = "查核人員：";

            run13.Append(runProperties13);
            run13.Append(text13);

            paragraph9.Append(paragraphProperties9);
            paragraph9.Append(run13);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph9);

            tableRow3.Append(tableCell7);
            tableRow3.Append(tableCell8);
            tableRow3.Append(tableCell9);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);

            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties10 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties10 = new ParagraphMarkRunProperties();
            RunFonts runFonts23 = new RunFonts() { Hint = FontTypeHintValues.EastAsia };

            paragraphMarkRunProperties10.Append(runFonts23);

            paragraphProperties10.Append(paragraphMarkRunProperties10);

            paragraph10.Append(paragraphProperties10);

            Table table2 = new Table();

            TableProperties tableProperties2 = new TableProperties();
            TableWidth tableWidth2 = new TableWidth() { Width = "19300", Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook2 = new TableLook() { Val = "04A0" };

            tableProperties2.Append(tableWidth2);
            tableProperties2.Append(tableBorders1);
            tableProperties2.Append(tableLook2);

            TableGrid tableGrid2 = new TableGrid();
            GridColumn gridColumn4 = new GridColumn() { Width = "1300" };
            GridColumn gridColumn5 = new GridColumn() { Width = "10000" };
            GridColumn gridColumn6 = new GridColumn() { Width = "3000" };
            GridColumn gridColumn7 = new GridColumn() { Width = "5000" };

            tableGrid2.Append(gridColumn4);
            tableGrid2.Append(gridColumn5);
            tableGrid2.Append(gridColumn6);
            tableGrid2.Append(gridColumn7);

            TableRow tableRow4 = new TableRow() { RsidTableRowMarkRevision = "00591921", RsidTableRowAddition = "008473F6", RsidTableRowProperties = "00591921" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableHeader tableHeader1 = new TableHeader();

            tableRowProperties1.Append(tableHeader1);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "1300", Type = TableWidthUnitValues.Dxa };
            Shading shading10 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties10.Append(tableCellWidth10);
            tableCellProperties10.Append(shading10);

            Paragraph paragraph11 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties11 = new ParagraphProperties();
            Justification justification3 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties11 = new ParagraphMarkRunProperties();
            RunFonts runFonts24 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties11.Append(runFonts24);

            paragraphProperties11.Append(justification3);
            paragraphProperties11.Append(paragraphMarkRunProperties11);

            Run run14 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties14 = new RunProperties();
            RunFonts runFonts25 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties14.Append(runFonts25);
            Text text14 = new Text();
            text14.Text = "編";

            run14.Append(runProperties14);
            run14.Append(text14);

            Run run15 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties15 = new RunProperties();
            RunFonts runFonts26 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties15.Append(runFonts26);
            Text text15 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text15.Text = " ";

            run15.Append(runProperties15);
            run15.Append(text15);

            Run run16 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties16 = new RunProperties();
            RunFonts runFonts27 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties16.Append(runFonts27);
            Text text16 = new Text();
            text16.Text = "號";

            run16.Append(runProperties16);
            run16.Append(text16);

            paragraph11.Append(paragraphProperties11);
            paragraph11.Append(run14);
            paragraph11.Append(run15);
            paragraph11.Append(run16);

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph11);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "10000", Type = TableWidthUnitValues.Dxa };
            Shading shading11 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties11.Append(tableCellWidth11);
            tableCellProperties11.Append(shading11);

            Paragraph paragraph12 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties12 = new ParagraphProperties();
            Justification justification4 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties12 = new ParagraphMarkRunProperties();
            RunFonts runFonts28 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties12.Append(runFonts28);

            paragraphProperties12.Append(justification4);
            paragraphProperties12.Append(paragraphMarkRunProperties12);

            Run run17 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties17 = new RunProperties();
            RunFonts runFonts29 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties17.Append(runFonts29);
            Text text17 = new Text();
            text17.Text = "查";

            run17.Append(runProperties17);
            run17.Append(text17);

            Run run18 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties18 = new RunProperties();
            RunFonts runFonts30 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties18.Append(runFonts30);
            Text text18 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text18.Text = "            ";

            run18.Append(runProperties18);
            run18.Append(text18);

            Run run19 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties19 = new RunProperties();
            RunFonts runFonts31 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties19.Append(runFonts31);
            Text text19 = new Text();
            text19.Text = "核";

            run19.Append(runProperties19);
            run19.Append(text19);

            Run run20 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties20 = new RunProperties();
            RunFonts runFonts32 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties20.Append(runFonts32);
            Text text20 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text20.Text = "            ";

            run20.Append(runProperties20);
            run20.Append(text20);

            Run run21 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties21 = new RunProperties();
            RunFonts runFonts33 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties21.Append(runFonts33);
            Text text21 = new Text();
            text21.Text = "項";

            run21.Append(runProperties21);
            run21.Append(text21);

            Run run22 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties22 = new RunProperties();
            RunFonts runFonts34 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties22.Append(runFonts34);
            Text text22 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text22.Text = "            ";

            run22.Append(runProperties22);
            run22.Append(text22);

            Run run23 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties23 = new RunProperties();
            RunFonts runFonts35 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties23.Append(runFonts35);
            Text text23 = new Text();
            text23.Text = "目";

            run23.Append(runProperties23);
            run23.Append(text23);

            paragraph12.Append(paragraphProperties12);
            paragraph12.Append(run17);
            paragraph12.Append(run18);
            paragraph12.Append(run19);
            paragraph12.Append(run20);
            paragraph12.Append(run21);
            paragraph12.Append(run22);
            paragraph12.Append(run23);

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph12);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "3000", Type = TableWidthUnitValues.Dxa };
            Shading shading12 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties12.Append(tableCellWidth12);
            tableCellProperties12.Append(shading12);

            Paragraph paragraph13 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties13 = new ParagraphProperties();
            Justification justification5 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties13 = new ParagraphMarkRunProperties();
            RunFonts runFonts36 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties13.Append(runFonts36);

            paragraphProperties13.Append(justification5);
            paragraphProperties13.Append(paragraphMarkRunProperties13);

            Run run24 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties24 = new RunProperties();
            RunFonts runFonts37 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties24.Append(runFonts37);
            Text text24 = new Text();
            text24.Text = "抽";

            run24.Append(runProperties24);
            run24.Append(text24);

            Run run25 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties25 = new RunProperties();
            RunFonts runFonts38 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties25.Append(runFonts38);
            Text text25 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text25.Text = " ";

            run25.Append(runProperties25);
            run25.Append(text25);

            Run run26 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties26 = new RunProperties();
            RunFonts runFonts39 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties26.Append(runFonts39);
            Text text26 = new Text();
            text26.Text = "查";

            run26.Append(runProperties26);
            run26.Append(text26);

            Run run27 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties27 = new RunProperties();
            RunFonts runFonts40 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties27.Append(runFonts40);
            Text text27 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text27.Text = " ";

            run27.Append(runProperties27);
            run27.Append(text27);

            Run run28 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties28 = new RunProperties();
            RunFonts runFonts41 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties28.Append(runFonts41);
            Text text28 = new Text();
            text28.Text = "範";

            run28.Append(runProperties28);
            run28.Append(text28);

            Run run29 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties29 = new RunProperties();
            RunFonts runFonts42 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties29.Append(runFonts42);
            Text text29 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text29.Text = " ";

            run29.Append(runProperties29);
            run29.Append(text29);

            Run run30 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties30 = new RunProperties();
            RunFonts runFonts43 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties30.Append(runFonts43);
            Text text30 = new Text();
            text30.Text = "圍";

            run30.Append(runProperties30);
            run30.Append(text30);

            paragraph13.Append(paragraphProperties13);
            paragraph13.Append(run24);
            paragraph13.Append(run25);
            paragraph13.Append(run26);
            paragraph13.Append(run27);
            paragraph13.Append(run28);
            paragraph13.Append(run29);
            paragraph13.Append(run30);

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph13);

            TableCell tableCell13 = new TableCell();

            TableCellProperties tableCellProperties13 = new TableCellProperties();
            TableCellWidth tableCellWidth13 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading13 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties13.Append(tableCellWidth13);
            tableCellProperties13.Append(shading13);

            Paragraph paragraph14 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties14 = new ParagraphProperties();
            Justification justification6 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties14 = new ParagraphMarkRunProperties();
            RunFonts runFonts44 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties14.Append(runFonts44);

            paragraphProperties14.Append(justification6);
            paragraphProperties14.Append(paragraphMarkRunProperties14);

            Run run31 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties31 = new RunProperties();
            RunFonts runFonts45 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties31.Append(runFonts45);
            Text text31 = new Text();
            text31.Text = "查";

            run31.Append(runProperties31);
            run31.Append(text31);

            Run run32 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties32 = new RunProperties();
            RunFonts runFonts46 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties32.Append(runFonts46);
            Text text32 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text32.Text = "    ";

            run32.Append(runProperties32);
            run32.Append(text32);

            Run run33 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties33 = new RunProperties();
            RunFonts runFonts47 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties33.Append(runFonts47);
            Text text33 = new Text();
            text33.Text = "核";

            run33.Append(runProperties33);
            run33.Append(text33);

            Run run34 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties34 = new RunProperties();
            RunFonts runFonts48 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties34.Append(runFonts48);
            Text text34 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text34.Text = "    ";

            run34.Append(runProperties34);
            run34.Append(text34);

            Run run35 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties35 = new RunProperties();
            RunFonts runFonts49 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties35.Append(runFonts49);
            Text text35 = new Text();
            text35.Text = "意";

            run35.Append(runProperties35);
            run35.Append(text35);

            Run run36 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties36 = new RunProperties();
            RunFonts runFonts50 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties36.Append(runFonts50);
            Text text36 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text36.Text = "    ";

            run36.Append(runProperties36);
            run36.Append(text36);

            Run run37 = new Run() { RsidRunProperties = "00591921" };

            RunProperties runProperties37 = new RunProperties();
            RunFonts runFonts51 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties37.Append(runFonts51);
            Text text37 = new Text();
            text37.Text = "見";

            run37.Append(runProperties37);
            run37.Append(text37);

            paragraph14.Append(paragraphProperties14);
            paragraph14.Append(run31);
            paragraph14.Append(run32);
            paragraph14.Append(run33);
            paragraph14.Append(run34);
            paragraph14.Append(run35);
            paragraph14.Append(run36);
            paragraph14.Append(run37);

            tableCell13.Append(tableCellProperties13);
            tableCell13.Append(paragraph14);

            tableRow4.Append(tableRowProperties1);
            tableRow4.Append(tableCell10);
            tableRow4.Append(tableCell11);
            tableRow4.Append(tableCell12);
            tableRow4.Append(tableCell13);

            table2.Append(tableProperties2);
            table2.Append(tableGrid2);
            table2.Append(tableRow4);
            Dictionary<string, string> dic = new Dictionary<string, string>();
            dic.Add("@plantypeid", Row_L1["PID"].ToString());
            dic.Add("@GID", Row_L1["GID"].ToString());
            dic.Add("@working_paperid", Row_L1["working_paperid"].ToString());
            dt_L2 = RptTool.ExecSqlQueryParameters(connectionString, @"select dbo.Auditfn_GetLang_workingpaperattach_list((select TOP(1)attachid from audit_system_auditplan_workingpaper_itemfield_attachments
							where audit_system_auditplan_workingpaper_itemfield_attachments.fieldid in(
							select audit_system_auditplan_workingpaper_opinions_itemfield.guidid from audit_system_auditplan_workingpaper_opinions_itemfield 
                            where audit_system_auditplan_workingpaper_opinions_itemfield.opinionsid in(
                            select guidid from audit_system_auditplan_workingpaper_opinions where audit_system_auditplan_workingpaper_opinions.itemid= item.guidid 
                            and audit_system_auditplan_workingpaper_opinions.ori_itemid= item.ori_guid
                            )
							)
							),'zh-tw') AS AttName
                            ,list.audit_no AS Numbers,(select itemvalue+',' from audit_system_auditplan_workingpaper_opinions_itemfield 
                            where audit_system_auditplan_workingpaper_opinions_itemfield.opinionsid in(
                            select guidid from audit_system_auditplan_workingpaper_opinions where audit_system_auditplan_workingpaper_opinions.itemid= item.guidid 
                            and audit_system_auditplan_workingpaper_opinions.ori_itemid= item.ori_guid
                            ) for xml path('')) as Content,*,
                            (case when audit_range='prevAudit'then '前次檢查至本次查核基準日'  ELSE  '本次查核首日' END)AS Range,
                            dbo.Auditfn_GetLang_plantype_list(list.plantypeid,'zh-tw') AS plantypeName
                            from audit_system_auditplan_list list 
                            left join audit_system_auditplan_assignments assign
                            on list.guidid=assign.planid
                            left join audit_system_auditplan_workingpaper_list wlist
                            on wlist.ori_guid=assign.working_paperid and wlist.planid=assign.planid
                            left join audit_system_auditplan_workingpaper_item item
                            on item.working_paperid=wlist.guidid
                            left join audit_system_auditplan_depts depts
                            on depts.planid=assign.planid
                                            where list.plantypeid = @plantypeid
                                            and(
                                            (select responsible_type from audit_org_dept_subgroup_list where guidid = '') = 'all'
                                                or
                                            (depts.deptid in (
                                            select responsible_id from audit_org_dept_subgroup_responsible where groupid = @GID)
                                            and(select responsible_type from audit_org_dept_subgroup_list where guidid = @GID) <> 'all'
                                            )) and assign.working_paperid = @working_paperid", dic);
            for (int i = 0; i < dt_L2.Rows.Count; i++)
            {
                table2.Append(AddRow1(dt_L2.Rows[i]));
            }

            Paragraph paragraph23 = new Paragraph() { RsidParagraphAddition = "00B763C0", RsidRunAdditionDefault = "00B763C0" };
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "page1", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph23.Append(bookmarkStart1);
            paragraph23.Append(bookmarkEnd1);

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "00B763C0", RsidSect = "00342541" };
            FooterReference footerReference1 = new FooterReference() { Type = HeaderFooterValues.Default, Id = "rId6" };
            PageSize pageSize1 = new PageSize() { Width = (UInt32Value)20639U, Height = (UInt32Value)14572U, Orient = PageOrientationValues.Landscape, Code = (UInt16Value)12U };
            PageMargin pageMargin1 = new PageMargin() { Top = 720, Right = (UInt32Value)720U, Bottom = 720, Left = (UInt32Value)720U, Header = (UInt32Value)851U, Footer = (UInt32Value)992U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "425" };
            DocGrid docGrid1 = new DocGrid() { Type = DocGridValues.Lines, LinePitch = 360 };

            sectionProperties1.Append(footerReference1);
            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(table1);
            body1.Append(paragraph10);
            body1.Append(table2);
            body1.Append(paragraph23);
            body1.Append(sectionProperties1);

            if (RowIndex < dt_L1.Rows.Count-1) {
                body1.Append(RptTool.GetBreakValues());            
            }
            return body1;
        }

        TableRow AddRow1(DataRow Row_L2)
        {
            TableRow tableRow6 = new TableRow() { RsidTableRowMarkRevision = "00591921", RsidTableRowAddition = "008473F6", RsidTableRowProperties = "00591921" };

            TableCell tableCell18 = new TableCell();

            TableCellProperties tableCellProperties18 = new TableCellProperties();
            TableCellWidth tableCellWidth18 = new TableCellWidth() { Width = "1300", Type = TableWidthUnitValues.Dxa };
            Shading shading18 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties18.Append(tableCellWidth18);
            tableCellProperties18.Append(shading18);

            Paragraph paragraph19 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "000D0440" };

            ParagraphProperties paragraphProperties19 = new ParagraphProperties();
            Justification justification10 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties19 = new ParagraphMarkRunProperties();
            RunFonts runFonts61 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties19.Append(runFonts61);

            paragraphProperties19.Append(justification10);
            paragraphProperties19.Append(paragraphMarkRunProperties19);

            Run run43 = new Run();

            RunProperties runProperties43 = new RunProperties();
            RunFonts runFonts62 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties43.Append(runFonts62);
            Text text43 = new Text();
            text43.Text = Row_L2["Numbers"].ToString();

            run43.Append(runProperties43);
            run43.Append(text43);

            paragraph19.Append(paragraphProperties19);
            paragraph19.Append(run43);

            tableCell18.Append(tableCellProperties18);
            tableCell18.Append(paragraph19);

            TableCell tableCell19 = new TableCell();

            TableCellProperties tableCellProperties19 = new TableCellProperties();
            TableCellWidth tableCellWidth19 = new TableCellWidth() { Width = "10000", Type = TableWidthUnitValues.Dxa };
            Shading shading19 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties19.Append(tableCellWidth19);
            tableCellProperties19.Append(shading19);

            Paragraph paragraph20 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "008473F6", RsidRunAdditionDefault = "000D0440" };

            ParagraphProperties paragraphProperties20 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties20 = new ParagraphMarkRunProperties();
            RunFonts runFonts63 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties20.Append(runFonts63);

            paragraphProperties20.Append(paragraphMarkRunProperties20);

            Run run44 = new Run();

            RunProperties runProperties44 = new RunProperties();
            RunFonts runFonts64 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties44.Append(runFonts64);
            Text text44 = new Text();
            text44.Text = Row_L2["Content"].ToString();

            run44.Append(runProperties44);
            run44.Append(text44);

            paragraph20.Append(paragraphProperties20);
            paragraph20.Append(run44);

            tableCell19.Append(tableCellProperties19);
            tableCell19.Append(paragraph20);

            TableCell tableCell20 = new TableCell();

            TableCellProperties tableCellProperties20 = new TableCellProperties();
            TableCellWidth tableCellWidth20 = new TableCellWidth() { Width = "3000", Type = TableWidthUnitValues.Dxa };
            Shading shading20 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties20.Append(tableCellWidth20);
            tableCellProperties20.Append(shading20);

            Paragraph paragraph21 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "000D0440" };

            ParagraphProperties paragraphProperties21 = new ParagraphProperties();
            Justification justification11 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties21 = new ParagraphMarkRunProperties();
            RunFonts runFonts66 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties21.Append(runFonts66);

            paragraphProperties21.Append(justification11);
            paragraphProperties21.Append(paragraphMarkRunProperties21);

            Run run46 = new Run();

            RunProperties runProperties46 = new RunProperties();
            RunFonts runFonts67 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            runProperties46.Append(runFonts67);
            Text text46 = new Text();
            text46.Text = Row_L2["Range"].ToString();

            run46.Append(runProperties46);
            run46.Append(text46);
            
            paragraph21.Append(paragraphProperties21);
            paragraph21.Append(run46);

            tableCell20.Append(tableCellProperties20);
            tableCell20.Append(paragraph21);

            TableCell tableCell21 = new TableCell();

            TableCellProperties tableCellProperties21 = new TableCellProperties();
            TableCellWidth tableCellWidth21 = new TableCellWidth() { Width = "5000", Type = TableWidthUnitValues.Dxa };
            Shading shading21 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties21.Append(tableCellWidth21);
            tableCellProperties21.Append(shading21);

            Paragraph paragraph22 = new Paragraph() { RsidParagraphMarkRevision = "00591921", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "00591921", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties22 = new ParagraphProperties();
            Justification justification12 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties22 = new ParagraphMarkRunProperties();
            RunFonts runFonts69 = new RunFonts() { Hint = FontTypeHintValues.EastAsia, Ascii = "Times New Roman", HighAnsi = "Times New Roman", EastAsia = "標楷體", ComplexScript = "Times New Roman" };

            paragraphMarkRunProperties22.Append(runFonts69);

            paragraphProperties22.Append(justification12);
            paragraphProperties22.Append(paragraphMarkRunProperties22);

            paragraph22.Append(paragraphProperties22);

            tableCell21.Append(tableCellProperties21);
            tableCell21.Append(paragraph22);

            tableRow6.Append(tableCell18);
            tableRow6.Append(tableCell19);
            tableRow6.Append(tableCell20);
            tableRow6.Append(tableCell21);

            return tableRow6;
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

            Font font3 = new Font() { Name = "Arial" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020B0604020202020204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "E0002AFF", UnicodeSignature1 = "C0007841", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            Font font5 = new Font() { Name = "標楷體" };
            Panose1Number panose1Number5 = new Panose1Number() { Val = "03000509000000000000" };
            FontCharSet fontCharSet5 = new FontCharSet() { Val = "88" };
            FontFamily fontFamily5 = new FontFamily() { Val = FontFamilyValues.Script };
            Pitch pitch5 = new Pitch() { Val = FontPitchValues.Fixed };
            FontSignature fontSignature5 = new FontSignature() { UnicodeSignature0 = "00000003", UnicodeSignature1 = "080E0000", UnicodeSignature2 = "00000016", UnicodeSignature3 = "00000000", CodePageSignature0 = "00100001", CodePageSignature1 = "00000000" };

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
            Zoom zoom1 = new Zoom() { Percent = "80" };
            BordersDoNotSurroundHeader bordersDoNotSurroundHeader1 = new BordersDoNotSurroundHeader();
            BordersDoNotSurroundFooter bordersDoNotSurroundFooter1 = new BordersDoNotSurroundFooter();
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 480 };
            DrawingGridHorizontalSpacing drawingGridHorizontalSpacing1 = new DrawingGridHorizontalSpacing() { Val = "120" };
            DisplayHorizontalDrawingGrid displayHorizontalDrawingGrid1 = new DisplayHorizontalDrawingGrid() { Val = 0 };
            DisplayVerticalDrawingGrid displayVerticalDrawingGrid1 = new DisplayVerticalDrawingGrid() { Val = 2 };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.CompressPunctuation };

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
            ApplyBreakingRules applyBreakingRules1 = new ApplyBreakingRules();
            UseFarEastLayout useFarEastLayout1 = new UseFarEastLayout();

            compatibility1.Append(spaceForUnderline1);
            compatibility1.Append(balanceSingleByteDoubleByteWidth1);
            compatibility1.Append(doNotLeaveBackslashAlone1);
            compatibility1.Append(underlineTrailingSpaces1);
            compatibility1.Append(doNotExpandShiftReturn1);
            compatibility1.Append(adjustLineHeightInTable1);
            compatibility1.Append(applyBreakingRules1);
            compatibility1.Append(useFarEastLayout1);

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "007E498B" };
            Rsid rsid1 = new Rsid() { Val = "00001D87" };
            Rsid rsid2 = new Rsid() { Val = "000C7B5C" };
            Rsid rsid3 = new Rsid() { Val = "000D0440" };
            Rsid rsid4 = new Rsid() { Val = "00261D34" };
            Rsid rsid5 = new Rsid() { Val = "00342541" };
            Rsid rsid6 = new Rsid() { Val = "00591921" };
            Rsid rsid7 = new Rsid() { Val = "006D59C0" };
            Rsid rsid8 = new Rsid() { Val = "007E498B" };
            Rsid rsid9 = new Rsid() { Val = "008473F6" };
            Rsid rsid10 = new Rsid() { Val = "00A00E67" };
            Rsid rsid11 = new Rsid() { Val = "00B763C0" };
            Rsid rsid12 = new Rsid() { Val = "00BA75D7" };
            Rsid rsid13 = new Rsid() { Val = "00CF5F85" };
            Rsid rsid14 = new Rsid() { Val = "00E37807" };

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

            ShapeDefaults shapeDefaults2 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults3 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 3074 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults2.Append(shapeDefaults3);
            shapeDefaults2.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };

            settings1.Append(zoom1);
            settings1.Append(bordersDoNotSurroundHeader1);
            settings1.Append(bordersDoNotSurroundFooter1);
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
            RunFonts runFonts70 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", EastAsia = "新細明體", ComplexScript = "Arial" };
            Languages languages1 = new Languages() { Val = "en-US", EastAsia = "zh-TW", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts70);
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
            Rsid rsid15 = new Rsid() { Val = "00001D87" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            WidowControl widowControl1 = new WidowControl() { Val = false };

            styleParagraphProperties1.Append(widowControl1);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            Kern kern1 = new Kern() { Val = (UInt32Value)2U };
            FontSize fontSize23 = new FontSize() { Val = "24" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };

            styleRunProperties1.Append(kern1);
            styleRunProperties1.Append(fontSize23);
            styleRunProperties1.Append(fontSizeComplexScript1);

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid15);
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
            Rsid rsid16 = new Rsid() { Val = "00B763C0" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs1.Append(tabStop1);
            tabs1.Append(tabStop2);
            SnapToGrid snapToGrid1 = new SnapToGrid() { Val = false };

            styleParagraphProperties2.Append(tabs1);
            styleParagraphProperties2.Append(snapToGrid1);

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            FontSize fontSize24 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties2.Append(fontSize24);
            styleRunProperties2.Append(fontSizeComplexScript2);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(linkedStyle1);
            style5.Append(uIPriority4);
            style5.Append(semiHidden4);
            style5.Append(unhideWhenUsed4);
            style5.Append(rsid16);
            style5.Append(styleParagraphProperties2);
            style5.Append(styleRunProperties2);

            Style style6 = new Style() { Type = StyleValues.Character, StyleId = "a4", CustomStyle = true };
            StyleName styleName6 = new StyleName() { Val = "頁首 字元" };
            BasedOn basedOn2 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "a3" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Rsid rsid17 = new Rsid() { Val = "00B763C0" };

            StyleRunProperties styleRunProperties3 = new StyleRunProperties();
            Kern kern2 = new Kern() { Val = (UInt32Value)2U };

            styleRunProperties3.Append(kern2);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle2);
            style6.Append(uIPriority5);
            style6.Append(semiHidden5);
            style6.Append(rsid17);
            style6.Append(styleRunProperties3);

            Style style7 = new Style() { Type = StyleValues.Paragraph, StyleId = "a5" };
            StyleName styleName7 = new StyleName() { Val = "footer" };
            BasedOn basedOn3 = new BasedOn() { Val = "a" };
            LinkedStyle linkedStyle3 = new LinkedStyle() { Val = "a6" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden6 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed5 = new UnhideWhenUsed();
            Rsid rsid18 = new Rsid() { Val = "00B763C0" };

            StyleParagraphProperties styleParagraphProperties3 = new StyleParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Center, Position = 4153 };
            TabStop tabStop4 = new TabStop() { Val = TabStopValues.Right, Position = 8306 };

            tabs2.Append(tabStop3);
            tabs2.Append(tabStop4);
            SnapToGrid snapToGrid2 = new SnapToGrid() { Val = false };

            styleParagraphProperties3.Append(tabs2);
            styleParagraphProperties3.Append(snapToGrid2);

            StyleRunProperties styleRunProperties4 = new StyleRunProperties();
            FontSize fontSize25 = new FontSize() { Val = "20" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "20" };

            styleRunProperties4.Append(fontSize25);
            styleRunProperties4.Append(fontSizeComplexScript3);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(linkedStyle3);
            style7.Append(uIPriority6);
            style7.Append(semiHidden6);
            style7.Append(unhideWhenUsed5);
            style7.Append(rsid18);
            style7.Append(styleParagraphProperties3);
            style7.Append(styleRunProperties4);

            Style style8 = new Style() { Type = StyleValues.Character, StyleId = "a6", CustomStyle = true };
            StyleName styleName8 = new StyleName() { Val = "頁尾 字元" };
            BasedOn basedOn4 = new BasedOn() { Val = "a0" };
            LinkedStyle linkedStyle4 = new LinkedStyle() { Val = "a5" };
            UIPriority uIPriority7 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden7 = new SemiHidden();
            Rsid rsid19 = new Rsid() { Val = "00B763C0" };

            StyleRunProperties styleRunProperties5 = new StyleRunProperties();
            Kern kern3 = new Kern() { Val = (UInt32Value)2U };

            styleRunProperties5.Append(kern3);

            style8.Append(styleName8);
            style8.Append(basedOn4);
            style8.Append(linkedStyle4);
            style8.Append(uIPriority7);
            style8.Append(semiHidden7);
            style8.Append(rsid19);
            style8.Append(styleRunProperties5);

            Style style9 = new Style() { Type = StyleValues.Table, StyleId = "a7" };
            StyleName styleName9 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn5 = new BasedOn() { Val = "a1" };
            UIPriority uIPriority8 = new UIPriority() { Val = 59 };
            Rsid rsid20 = new Rsid() { Val = "008473F6" };

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders2 = new TableBorders();
            TopBorder topBorder2 = new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder2 = new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder2 = new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder2 = new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder2 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder2 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders2.Append(topBorder2);
            tableBorders2.Append(leftBorder2);
            tableBorders2.Append(bottomBorder2);
            tableBorders2.Append(rightBorder2);
            tableBorders2.Append(insideHorizontalBorder2);
            tableBorders2.Append(insideVerticalBorder2);

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin2);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableBorders2);
            styleTableProperties2.Append(tableCellMarginDefault2);

            style9.Append(styleName9);
            style9.Append(basedOn5);
            style9.Append(uIPriority8);
            style9.Append(rsid20);
            style9.Append(styleTableProperties2);

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

            styleDefinitionsPart1.Styles = styles1;
        }

        // Generates content of footerPart1.
        private void GenerateFooterPart1Content(FooterPart footerPart1)
        {
            Footer footer1 = new Footer();
            footer1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Paragraph paragraph24 = new Paragraph() { RsidParagraphMarkRevision = "008473F6", RsidParagraphAddition = "008473F6", RsidParagraphProperties = "008473F6", RsidRunAdditionDefault = "008473F6" };

            ParagraphProperties paragraphProperties23 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "a5" };
            Justification justification13 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties23 = new ParagraphMarkRunProperties();
            FontSize fontSize26 = new FontSize() { Val = "24" };

            paragraphMarkRunProperties23.Append(fontSize26);

            paragraphProperties23.Append(paragraphStyleId1);
            paragraphProperties23.Append(justification13);
            paragraphProperties23.Append(paragraphMarkRunProperties23);

            Run run48 = new Run() { RsidRunProperties = "008473F6" };

            RunProperties runProperties48 = new RunProperties();
            FontSize fontSize27 = new FontSize() { Val = "24" };

            runProperties48.Append(fontSize27);
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run48.Append(runProperties48);
            run48.Append(fieldChar1);

            Run run49 = new Run() { RsidRunProperties = "008473F6" };

            RunProperties runProperties49 = new RunProperties();
            FontSize fontSize28 = new FontSize() { Val = "24" };

            runProperties49.Append(fontSize28);
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " Page \\* MERGEFORMAT ";

            run49.Append(runProperties49);
            run49.Append(fieldCode1);

            Run run50 = new Run() { RsidRunProperties = "008473F6" };

            RunProperties runProperties50 = new RunProperties();
            FontSize fontSize29 = new FontSize() { Val = "24" };

            runProperties50.Append(fontSize29);
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run50.Append(runProperties50);
            run50.Append(fieldChar2);

            Run run51 = new Run() { RsidRunAddition = "00CF5F85" };

            RunProperties runProperties51 = new RunProperties();
            NoProof noProof1 = new NoProof();
            FontSize fontSize30 = new FontSize() { Val = "24" };

            runProperties51.Append(noProof1);
            runProperties51.Append(fontSize30);
            Text text48 = new Text();
            text48.Text = "1";

            run51.Append(runProperties51);
            run51.Append(text48);

            Run run52 = new Run() { RsidRunProperties = "008473F6" };

            RunProperties runProperties52 = new RunProperties();
            FontSize fontSize31 = new FontSize() { Val = "24" };

            runProperties52.Append(fontSize31);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run52.Append(runProperties52);
            run52.Append(fieldChar3);

            Run run53 = new Run() { RsidRunProperties = "008473F6" };

            RunProperties runProperties53 = new RunProperties();
            FontSize fontSize32 = new FontSize() { Val = "24" };

            runProperties53.Append(fontSize32);
            Text text49 = new Text();
            text49.Text = "/";

            run53.Append(runProperties53);
            run53.Append(text49);

            SimpleField simpleField1 = new SimpleField() { Instruction = " NumPages \\* MERGEFORMAT " };

            Run run54 = new Run() { RsidRunAddition = "00CF5F85" };

            RunProperties runProperties54 = new RunProperties();
            NoProof noProof2 = new NoProof();
            FontSize fontSize33 = new FontSize() { Val = "24" };

            runProperties54.Append(noProof2);
            runProperties54.Append(fontSize33);
            Text text50 = new Text();
            text50.Text = "1";

            run54.Append(runProperties54);
            run54.Append(text50);

            simpleField1.Append(run54);

            paragraph24.Append(paragraphProperties23);
            paragraph24.Append(run48);
            paragraph24.Append(run49);
            paragraph24.Append(run50);
            paragraph24.Append(run51);
            paragraph24.Append(run52);
            paragraph24.Append(run53);
            paragraph24.Append(simpleField1);

            footer1.Append(paragraph24);

            footerPart1.Footer = footer1;
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

            Paragraph paragraph25 = new Paragraph() { RsidParagraphAddition = "00E37807", RsidParagraphProperties = "00B763C0", RsidRunAdditionDefault = "00E37807" };

            Run run55 = new Run();
            SeparatorMark separatorMark1 = new SeparatorMark();

            run55.Append(separatorMark1);

            paragraph25.Append(run55);

            endnote1.Append(paragraph25);

            Endnote endnote2 = new Endnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph26 = new Paragraph() { RsidParagraphAddition = "00E37807", RsidParagraphProperties = "00B763C0", RsidRunAdditionDefault = "00E37807" };

            Run run56 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark1 = new ContinuationSeparatorMark();

            run56.Append(continuationSeparatorMark1);

            paragraph26.Append(run56);

            endnote2.Append(paragraph26);

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

            Paragraph paragraph27 = new Paragraph() { RsidParagraphAddition = "00E37807", RsidParagraphProperties = "00B763C0", RsidRunAdditionDefault = "00E37807" };

            Run run57 = new Run();
            SeparatorMark separatorMark2 = new SeparatorMark();

            run57.Append(separatorMark2);

            paragraph27.Append(run57);

            footnote1.Append(paragraph27);

            Footnote footnote2 = new Footnote() { Type = FootnoteEndnoteValues.ContinuationSeparator, Id = 0 };

            Paragraph paragraph28 = new Paragraph() { RsidParagraphAddition = "00E37807", RsidParagraphProperties = "00B763C0", RsidRunAdditionDefault = "00E37807" };

            Run run58 = new Run();
            ContinuationSeparatorMark continuationSeparatorMark2 = new ContinuationSeparatorMark();

            run58.Append(continuationSeparatorMark2);

            paragraph28.Append(run58);

            footnote2.Append(paragraph28);

            footnotes1.Append(footnote1);
            footnotes1.Append(footnote2);

            footnotesPart1.Footnotes = footnotes1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Rina";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2013-12-26T12:41:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2013-12-26T12:41:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Alan";
        }
    }
}
