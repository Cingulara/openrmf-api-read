// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace openrmf_read_api.Models
{
    public static class ExcelStyleSheet 
    {
        public static Stylesheet GenerateStylesheet()
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)3U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 14D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 18D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold2);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 30D };
            Color color3 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(bold1);
            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontScheme3);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            // 0 no BG color just normal Excel
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);

            // 1 red or open CAT 1
            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor2 = new ForegroundColor() { Rgb = "FFFF0000" };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill2.Append(foregroundColor2);
            patternFill2.Append(backgroundColor2);
            fill2.Append(patternFill2);

            // 2 silver or N/A
            Fill fill3 = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor3 = new ForegroundColor() { Rgb = "FFCCCCCC" };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor3);
            patternFill3.Append(backgroundColor3);
            fill3.Append(patternFill3);

            // 3 green or NaF
            Fill fill4 = new Fill();
            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor4 = new ForegroundColor() { Rgb = "FF50CC83" };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill4.Append(foregroundColor4);
            patternFill4.Append(backgroundColor4);
            fill4.Append(patternFill4);

            // 4 black&white or Not Reviewed
            Fill fill5 = new Fill();
            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor5 = new ForegroundColor() { Rgb = "FFFFFFFF" };
            BackgroundColor backgroundColor5 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill5.Append(foregroundColor5);
            patternFill5.Append(backgroundColor5);
            fill5.Append(patternFill5);

            // 5 red or open CAT 1
            Fill fill6 = new Fill();
            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor6 = new ForegroundColor() { Rgb = "FFE53935" };
            BackgroundColor backgroundColor6 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill6.Append(foregroundColor6);
            patternFill6.Append(backgroundColor6);
            fill6.Append(patternFill6);

            // 6 orange open CAT 2
            Fill fill7 = new Fill();
            PatternFill patternFill7 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor7 = new ForegroundColor() { Rgb = "FFFFA500" };
            BackgroundColor backgroundColor7 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill7.Append(foregroundColor7);
            patternFill7.Append(backgroundColor7);
            fill7.Append(patternFill7);

            // 7 orange open CAT 3
            Fill fill8 = new Fill();
            PatternFill patternFill8 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor8 = new ForegroundColor() { Rgb = "FFD8D80E" };
            BackgroundColor backgroundColor8 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill8.Append(foregroundColor8);
            patternFill8.Append(backgroundColor8);
            fill8.Append(patternFill8);

            // 8 gray background for any general use
            Fill fill9 = new Fill();
            PatternFill patternFill9 = new PatternFill() { PatternType = PatternValues.Solid};
            ForegroundColor foregroundColor9 = new ForegroundColor() { Rgb = "FFBBBBBB" };
            BackgroundColor backgroundColor9 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill9.Append(foregroundColor9);
            patternFill9.Append(backgroundColor9);
            fill9.Append(patternFill9);

            fills1.Append(fill1); // these are 0 based when you include in the CellFormat statements below, so 1 off
            fills1.Append(fill2);
            fills1.Append(fill3);
            fills1.Append(fill4);
            fills1.Append(fill5);
            fills1.Append(fill6);
            fills1.Append(fill7);
            fills1.Append(fill8);
            fills1.Append(fill9);

            Borders borders1 = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);
            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            cellStyleFormats1.Append(cellFormat1);
            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)4U };

            // style index 0:  normal font and wrapping of text for cell rows
            CellFormat cellFormat2 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 1:  normal font with numerical format
            CellFormat cellFormat3 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)2U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            // style index 2:  title font of 30 bold
            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            // style index 3:  info row under title and header rows font bold size 18
            CellFormat cellFormat5 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };

            // fill colors based on the fillx variables above to match the 4 statuses of the checklist vulnerabilities
            // style index 4: red or Open
            CellFormat cellFormat6 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 5: silver or Not Applicable
            CellFormat cellFormat7 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 6: green or Not a Finding
            CellFormat cellFormat8 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 7: b/w or Not Reviewed
            CellFormat cellFormat9 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            // CENTER The content for listings versus the excel export of a checklist above
            // style index 8: red or Open centered content
            CellFormat cellFormat10 = new CellFormat(new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 9: silver or Not Applicable
            CellFormat cellFormat11= new CellFormat(new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 10: green or Not a Finding
            CellFormat cellFormat12 = new CellFormat(new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 11: b/w or Not Reviewed
            CellFormat cellFormat13 = new CellFormat(new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            // style index 12: orange Open CAT 2
            CellFormat cellFormat14 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 13: yellow Open CAT 3
            CellFormat cellFormat15 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)7U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 14: orange Open CAT 2 centered
            CellFormat cellFormat16 = new CellFormat(new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)6U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 15: yellow Open CAT 3 centered
            CellFormat cellFormat17 = new CellFormat(new Alignment() { WrapText = true, Horizontal = HorizontalAlignmentValues.Center }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)7U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            // style index 16: yellow Open CAT 3 centered
            CellFormat cellFormat18 = new CellFormat(new Alignment() { WrapText = true }) { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)8U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };

            // add all these formats
            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);
            cellFormats1.Append(cellFormat8);
            cellFormats1.Append(cellFormat9);
            cellFormats1.Append(cellFormat10);
            cellFormats1.Append(cellFormat11);
            cellFormats1.Append(cellFormat12);
            cellFormats1.Append(cellFormat13);
            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);
            cellFormats1.Append(cellFormat18);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = 0, BuiltinId = 0 };
            cellStyles1.Append(cellStyle1);

            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = 0 };
            TableStyles tableStyles1 = new TableStyles() { Count = 0, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);
            return stylesheet1;
        }
    }
}