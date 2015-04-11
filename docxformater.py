# -*- coding: utf-8 -*-
from os import path
import re
import win32com.client
from win32com.client import constants as C

FigureLabel = u"图"
NormalFont = u"微软雅黑"
TableStyle = u"网格型"
NormalStyleName = u"正文"
CodeStyleName = "Source Code"
SourceCodeFont = "Source Code Pro"


class Converter(object):
    def __init__(self, factor, divisor=1):
        if factor != 1 and divisor != 1:
            self.factor = float(factor)
            self.divisor = float(divisor)
            self.operation = self.scale
        elif divisor != 1:
            self.divisor = float(divisor)
            self.operation = self.divide
        elif factor != 1:
            self.factor = float(factor)
            self.operation = self.multiply
        else:
            self.operation = self.noop

    def multiply(self, arg):
        return arg * self.factor

    def divide(self, arg):
        return arg / self.divisor

    def scale(self, arg):
        return arg * self.factor / self.divisor

    def noop(self, arg):
        return arg + 0.0

    def __call__(self, arg):
        return self.operation(arg)


PicasToPoints = Converter(12)
PointsToPicas = Converter(1, 12)
InchesToPoints = Converter(72)
PointsToInches = Converter(1, 72)
LinesToPoints = Converter(12)
PointsToLines = Converter(1, 12)
InchesToCentimeters = Converter(254, 100)
CentimetersToInches = Converter(100, 254)
CentimetersToPoints = Converter(InchesToPoints(100), 254)
PointsToCentimeters = Converter(254, InchesToPoints(100))


def goto_start():
    app.Selection.GoTo(What=c.wdGoToSection, Which=c.wdGoToFirst)


def set_block_style(start, end):
    app.Selection.Start = start
    app.Selection.End = end
    borders = [c.wdBorderTop, c.wdBorderBottom, c.wdBorderLeft, c.wdBorderRight]

    for border_id in borders:
        border = app.Selection.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.LineWidth = c.wdLineWidth050pt
        border.Color = 0xC0C0C0

    app.Selection.ParagraphFormat.Shading.BackgroundPatternColor = 0xE0E0E0
    b = app.Selection.ParagraphFormat.Borders
    b.DistanceFromTop = 4
    b.DistanceFromLeft = 4
    b.DistanceFromBottom = 4
    b.DistanceFromRight = 4


def set_code_style(paragraph, left_width, fill_color):
    paragraph.Range.Select()
    borders = [c.wdBorderTop, c.wdBorderBottom]

    for border_id in borders:
        border = app.Selection.Borders(border_id)
        border.LineStyle = c.wdLineStyleSingle
        border.LineWidth = c.wdLineWidth050pt
        border.Color = fill_color

    border = app.Selection.Borders(c.wdBorderLeft)
    border.LineStyle = c.wdLineStyleSingle
    border.LineWidth = left_width
    border.Color = 0xAFAFAF

    app.Selection.ParagraphFormat.Shading.BackgroundPatternColor = fill_color
    b = app.Selection.ParagraphFormat.Borders
    b.DistanceFromTop = 4
    b.DistanceFromLeft = 4
    b.DistanceFromBottom = 4
    b.DistanceFromRight = 4


def remove_first_line(paragraph):
    text = paragraph.Range.Text
    idx = text.index("\x0b")
    r = paragraph.Range
    r.End = r.Start + idx + 1
    r.Delete()


def add_graph_caption(shape):
    this_p = shape.Range.Paragraphs(1)
    next_p = this_p.Next()
    title = next_p.Range.Text
    next_p.Range.Delete()
    shape.Range.Select()
    app.Selection.InsertCaption(Label=FigureLabel, TitleAutoText=u"abc", Title=u"  " + title[:-1],
                                Position=c.wdCaptionPositionBelow, ExcludeLabel=0)


def get_graph_title(shape):
    this_p = shape.Range.Paragraphs(1)
    next_p = this_p.Next()
    title = next_p.Range.Text
    res = re.match(FigureLabel + u" \d+.(\d+)(.*)", title[:-1])
    return res.group(1), res.group(2)


def doc_find(pattern):
    find = app.Selection.Find
    find.MatchCase = False
    find.MatchWholeWord = False
    find.MatchAllWordForms = False
    find.MatchSoundsLike = False
    find.MatchWildcards = False
    find.MatchFuzzy = False
    res = find.Execute(pattern)
    return res


def process_all_codes():
    goto_start()
    while True:
        if doc_find("##CODE"):
            p = app.Selection.Range.Paragraphs(1)
            remove_first_line(p)
            set_code_style(p, c.wdLineWidth300pt, 0xF0F0F0)
            app.Selection.Move()
        else:
            break

    goto_start()

    while True:
        if doc_find("##OUTPUT"):
            p = app.Selection.Range.Paragraphs(1)
            remove_first_line(p)
            set_code_style(p, c.wdLineWidth100pt, 0xF7F7F7)
            app.Selection.Move()
        else:
            break


def process_all_graphs():
    goto_start()
    for shape in doc.InlineShapes:
        add_graph_caption(shape)


def process_all_references():
    import bisect

    shapes = list(doc.InlineShapes)
    shape_starts = [shape.Range.Start for shape in shapes]
    gids = [get_graph_title(shape)[0] for shape in shapes]

    def _process(target, offset):
        goto_start()
        reference_gids = []
        while True:
            if doc_find(target):
                r = app.Selection.Range
                start = r.Start
                end = r.End
                idx = bisect.bisect_left(shape_starts, start) + offset
                reference_gids.append(gids[idx])
            else:
                break

        goto_start()
        i = 0
        while True:
            if doc_find(target):
                gid = reference_gids[i]
                app.Selection.InsertCrossReference(ReferenceType=FigureLabel,
                                                   ReferenceKind=c.wdOnlyLabelAndNumber,
                                                   ReferenceItem=gid,
                                                   InsertAsHyperlink=True,
                                                   IncludePosition=False,
                                                   SeparateNumbers=False,
                                                   SeparatorString=" ")
                i += 1
            else:
                break

    _process("ref:fig-next", 0)
    for i in range(9, 1, -1):
        _process("ref:fig-prev{}".format(i), -i)
    _process("ref:fig-prev", -1)


def process_all_table():
    for table in doc.Tables:
        table.Style = TableStyle


def process_heading():
    list_temp = app.ListGalleries(c.wdOutlineNumberGallery).ListTemplates(1)
    list_temp.Name = ""
    goto_start()
    while True:
        app.Selection.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=list_temp,
            ContinuePreviousList=True,
            ApplyTo=c.wdListApplyToWholeList,
            DefaultListBehavior=c.wdWord10ListBehavior)
        start = app.Selection.Start
        r = app.Selection.GoTo(What=c.wdGoToHeading, Which=c.wdGoToNext, Count=1, Name="")
        if r.Start == start:
            break
    return


def process_blocks():
    goto_start()
    while True:
        res = doc_find("%BLOCK_START")
        if not res:
            break

        app.Selection.End += 1
        app.Selection.Delete()
        start = app.Selection.Start

        doc_find("%BLOCK_END")
        app.Selection.End += 1
        app.Selection.Delete()
        end = app.Selection.Start

        set_block_style(start, end)

        app.Selection.Move()


def process_miniblock(btype, pcount, marker):
    goto_start()

    while True:
        res = doc_find("%{}".format(btype))
        if not res:
            break
        app.Selection.End += 1
        app.Selection.Delete()
        app.Selection.MoveDown(Unit=C.wdParagraph, Count=pcount, Extend=C.wdExtend)
        app.Selection.Cut()
        table = doc.Tables.Add(Range=app.Selection.Range, NumRows=1, NumColumns=2,
                               DefaultTableBehavior=C.wdWord9TableBehavior,
                               AutoFitBehavior=C.wdAutoFitFixed)
        cell1 = table.Cell(1, 1)
        cell2 = table.Cell(1, 2)

        for cell in (cell1, cell2):
            for bid in range(1, 5):
                cell.Borders(bid).LineStyle = C.wdLineStyleNone

        cell1.Range.Text = marker
        cell1.Range.Font.Size = 32
        cell1.Range.Font.Name = u"Segoe UI Symbol"
        cell1.Range.Font.Color = 0xA0A0A0
        cell1.Width = CentimetersToPoints(1.5)

        cell2 = table.Cell(1, 2)
        cell2.Range.Select()
        app.Selection.PasteAndFormat(C.wdFormatOriginalFormatting)
        app.Selection.TypeBackspace()
        cell2.Width = CentimetersToPoints(16)
        cell2.VerticalAlignment = C.wdCellAlignVerticalCenter

        table.Shading.BackgroundPatternColor = 0xF0F0F0
        for border_id in (C.wdBorderTop, C.wdBorderBottom):
            border = table.Borders(border_id)
            border.LineStyle = C.wdLineStyleSingle
            border.LineWidth = C.wdLineWidth050pt
            border.Color = 0xC0C0C0


def process_all_miniblocks():
    process_miniblock("LINK", 2, u"🌏")
    process_miniblock("WARNING", 1, u"⚠")
    process_miniblock("TIP", 1, u"💡")
    process_miniblock("QUESTION", 1, u"❓")
    process_miniblock("SOURCE", 1, u"📀")


def process_page():
    page = doc.PageSetup

    page.TopMargin = CentimetersToPoints(2.54)
    page.BottomMargin = CentimetersToPoints(2.54)
    page.LeftMargin = CentimetersToPoints(1.9)
    page.RightMargin = CentimetersToPoints(1.9)
    page.Gutter = CentimetersToPoints(0)
    page.HeaderDistance = CentimetersToPoints(1.27)
    page.FooterDistance = CentimetersToPoints(1.27)
    page.PageWidth = CentimetersToPoints(21)
    page.PageHeight = CentimetersToPoints(29.7)

    doc.Styles(NormalStyleName).Font.Size = 10.5
    doc.Styles(NormalStyleName).Font.Name = NormalFont

    pformat = doc.Styles(NormalStyleName).ParagraphFormat
    pformat.LineSpacingRule = C.wdLineSpaceMultiple
    pformat.LineSpacing = LinesToPoints(0.9)
    pformat.SpaceAfter = 6
    pformat.SpaceBefore = 6

    font = doc.Styles(CodeStyleName).Font
    # font.NameFarEast = "+中文正文"
    font.NameAscii = SourceCodeFont
    font.NameOther = SourceCodeFont
    font.Name = SourceCodeFont
    font.Size = 9

    pformat = doc.Styles(CodeStyleName).ParagraphFormat
    pformat.LineSpacing = LinesToPoints(0.9)

    for i in range(1, 6):
        level = app.ListGalleries(c.wdOutlineNumberGallery).ListTemplates(1).ListLevels(i)
        level.NumberFormat = ".".join("%" + str(j) for j in range(1, i + 1)) + "."
        level.TrailingCharacter = c.wdTrailingTab
        level.NumberStyle = c.wdListNumberStyleArabic
        level.NumberPosition = CentimetersToPoints(0)
        level.Alignment = c.wdListLevelAlignLeft
        level.TextPosition = CentimetersToPoints(0.75)
        level.TabPosition = c.wdUndefined
        if i == 1:
            print doc.Name
            basename = path.basename(doc.Name)
            level.StartAt = int(basename.split("-")[0])

    label = app.CaptionLabels(FigureLabel)
    label.NumberStyle = c.wdCaptionNumberStyleArabic
    label.IncludeChapterNumber = True
    label.ChapterStyleLevel = 1
    label.Separator = c.wdSeparatorHyphen


def process_docx(docx_fn):
    global app, doc, c
    from .common import BUILD_ROOT_FOLDER

    abs_docx_fn = path.abspath(path.join(BUILD_ROOT_FOLDER, "build_" + docx_fn, docx_fn) + ".docx")
    print abs_docx_fn

    app = win32com.client.gencache.EnsureDispatch("Word.Application")
    app.Visible = True
    from win32com.client import constants as c

    doc = app.Documents.Open(abs_docx_fn)

    process_page()
    process_all_codes()
    process_heading()
    process_all_graphs()
    process_all_references()
    process_all_table()
    process_all_miniblocks()
    process_blocks()


if __name__ == '__main__':
    process_docx("02-numpy")
