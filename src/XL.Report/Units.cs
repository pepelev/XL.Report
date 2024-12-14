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

using XL.Report.Contents;
using XL.Report.Styles;
using String = XL.Report.Contents.String;

namespace XL.Report;

public sealed class Units
{
    private readonly Book book;

    public Units(Book book)
    {
        this.book = book;
    }

    public Cell BlankCell(StyleId? styleId = null) => new(Content.Blank, styleId);
    public Cell BlankCell(Style style) => BlankCell(book.Styles.Register(style));

    public Cell Cell(Content content, StyleId? styleId = null) => new(content, styleId);
    public Cell Cell(Content content, Style style) => Cell(content, book.Styles.Register(style));

    public Cell Cell(string value, StyleId? styleId = null) => new(new String(value, book.Strings), styleId);
    public Cell Cell(string value, Style style) => Cell(value, book.Styles.Register(style));

    public Cell Cell(long value, StyleId? styleId = null) => new(new Number.Integral(value), styleId);
    public Cell Cell(long value, Style style) => Cell(value, book.Styles.Register(style));

    public Cell Cell(double value, StyleId? styleId = null) => new(new Number.Fractional(value), styleId);
    public Cell Cell(double value, Style style) => Cell(value, book.Styles.Register(style));

    public Cell Cell(decimal value, StyleId? styleId = null) => new(new Number.Financial(value), styleId);
    public Cell Cell(decimal value, Style style) => Cell(value, book.Styles.Register(style));

    public Cell Cell(DateTime value, StyleId styleId) => new(new Number.Instant(value), styleId);
    public Cell Cell(DateTime value, Style style) => Cell(value, book.Styles.Register(style));

    public Cell Cell(bool value, StyleId? styleId = null) => new(Bool.From(value), styleId);
    public Cell Cell(bool value, Style style) => Cell(value, book.Styles.Register(style));

    public Cell Cell(Expression value, StyleId? styleId = null) => new(new Formula(value), styleId);
    public Cell Cell(Expression value, Style style) => Cell(value, book.Styles.Register(style));

    public Merge BlankMerge(Size size, StyleId? styleId = null) => new(size, Content.Blank, styleId);
    public Merge BlankMerge(Size size, Style style) => BlankMerge(size, book.Styles.Register(style));

    public Merge Merge(Content content, Size size, StyleId? styleId = null) => new(size, content, styleId);
    public Merge Merge(Content content, Size size, Style style) => Merge(content, size, book.Styles.Register(style));

    public Merge Merge(string value, Size size, StyleId? styleId = null) => new(size, new String(value, book.Strings), styleId);
    public Merge Merge(string value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));

    public Merge Merge(long value, Size size, StyleId? styleId = null) => new(size, new Number.Integral(value), styleId);
    public Merge Merge(long value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));

    public Merge Merge(double value, Size size, StyleId? styleId = null) => new(size, new Number.Fractional(value), styleId);
    public Merge Merge(double value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));

    public Merge Merge(decimal value, Size size, StyleId? styleId = null) => new(size, new Number.Financial(value), styleId);
    public Merge Merge(decimal value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));

    public Merge Merge(DateTime value, Size size, StyleId styleId) => new(size, new Number.Instant(value), styleId);
    public Merge Merge(DateTime value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));

    public Merge Merge(bool value, Size size, StyleId? styleId = null) => new(size, Bool.From(value), styleId);
    public Merge Merge(bool value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));

    public Merge Merge(Expression value, Size size, StyleId? styleId = null) => new(size, new Formula(value), styleId);
    public Merge Merge(Expression value, Size size, Style style) => Merge(value, size, book.Styles.Register(style));
}