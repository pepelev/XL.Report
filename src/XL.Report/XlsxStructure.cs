#region Legal
// Copyright 2024 Pepelev Alexey
// 
// This file is part of XL.Report.
// 
// XL.Report is free software: you can redistribute it and/or modify it under the terms of the
// GNU Lesser General Public License as published by the Free Software Foundation, either
// version 3 of the License, or (at your option) any later version.
// 
// XL.Report is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
// See the GNU Lesser General Public License for more details.
// 
// You should have received a copy of the GNU Lesser General Public License along with XL.Report.
// If not, see <https://www.gnu.org/licenses/>.
#endregion

namespace XL.Report;

public static class XlsxStructure
{
    public static class Namespaces
    {
        public static class OfficeDocuments
        {
            public const string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        }

        public static class Spreadsheet
        {
            public const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        }
    }

    public static class Worksheet
    {
        public const string Row = "row";
        public const string Cell = "c";
        public const string Reference = "r";
        public const string Style = "s";
    }

    public static class Styles
    {
        public const string RootElement = "styleSheet";

        public static class Fonts
        {
            public const string Collection = "fonts";
            public const string CollectionCount = "count";
            public const string Font = "font";
            public const string Size = "sz";
            public const string SizeValue = "val";
            public const string Color = "color";
            public const string ColorRgb = "rgb";
            public const string Name = "name";
            public const string NameValue = "val";
            public const string Family = "family";
            public const string FamilyValue = "val";
            public const string Charset = "charset";
            public const string CharsetValue = "val";
            public const string Scheme = "minor";
            public const string SchemeValue = "val";
        }

        public static class Fills
        {
            public const string Collection = "fills";
            public const string Fill = "fill";
            public const string Pattern = "patternFill";
            public const string PatternType = "patternType";
        }

        public static class Borders
        {
            public const string Collection = "borders";
            public const string Border = "border";
            public const string Left = "left";
            public const string Right = "right";
            public const string Top = "top";
            public const string Bottom = "bottom";
            public const string Diagonal = "diagonal";
        }

        public static class CellFormats
        {
            public const string StyleFormats = "cellStyleXfs";
            public const string StyleFormatsCount = "count";
            public const string StyleFormatIndex = "xfId";

            public const string Collection = "cellXfs";
            public const string CollectionCount = "count";

            public const string CellFormat = "xf";
            public const string NumberFormatIndex = "numFmtId";
            public const string FontIndex = "fontId";
            public const string FillIndex = "fillId";
            public const string BordersIndex = "borderId";

            public const string ApplyFormat = "applyNumberFormat";
            public const string ApplyFont = "applyFont";
            public const string ApplyFill = "applyFill";
            public const string ApplyBorders = "applyBorder";
            public const string ApplyAlignment = "applyAlignment";
        }
    }
}