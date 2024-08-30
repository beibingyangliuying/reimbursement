from enum import Enum, auto
from types import MappingProxyType

import pandas as pd
from cytoolz.curried import curry, memoize  # type:ignore
from docx.document import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pymonad.reader import Pipe  # type:ignore

Date = str
Name = str
Category = str
HasInvoice = str
Description = str
Amount = float


class ColumnType(Enum):
    DATE = auto()
    NAME = auto()
    CATEGORY = auto()
    HASINVOICE = auto()
    DESCRIPTION = auto()
    AMOUNT = auto()


HEADINGS = MappingProxyType(
    {
        ColumnType.DATE: "日期",
        ColumnType.NAME: "姓名",
        ColumnType.CATEGORY: "类别",
        ColumnType.HASINVOICE: "是否有发票",
        ColumnType.DESCRIPTION: "说明",
        ColumnType.AMOUNT: "金额",
    }
)
COLORS = MappingProxyType(
    {
        "black": RGBColor(0, 0, 0),
        "red": RGBColor(255, 0, 0),
        "blue": RGBColor(0, 0, 255),
        "green": RGBColor(0, 255, 0),
    }
)
FONTSIZE = MappingProxyType(
    {
        "初号": Pt(42),
        "小初": Pt(36),
        "一号": Pt(26),
        "小一": Pt(24),
        "二号": Pt(22),
        "小二": Pt(18),
        "三号": Pt(16),
        "小三": Pt(15),
        "四号": Pt(14),
        "小四": Pt(12),
        "五号": Pt(10.5),
        "小五": Pt(9),
        "六号": Pt(7.5),
        "小六": Pt(6.5),
        "七号": Pt(5.5),
        "八号": Pt(5),
    }
)


class FontFamily(Enum):
    ROMAN = auto()
    ITALIC = auto()
    BOLD = auto()


def init_blank_document() -> Document:
    def func(style, font_family: FontFamily) -> None:  # type:ignore
        match font_family:
            case FontFamily.ROMAN:
                style.font.name = "Times New Roman"
                style.font._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
            case FontFamily.ITALIC:
                style.font.name = "Times New Roman"
                style.font._element.rPr.rFonts.set(qn("w:eastAsia"), "楷体")
            case FontFamily.BOLD:
                style.font.name = "Arial"
                style.font._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")
        # May raise AttributeError

    from docx import Document as create_document

    doc = create_document()

    style = doc.styles["Normal"]
    func(style, FontFamily.ROMAN)

    style = doc.styles["Emphasis"]
    func(style, FontFamily.ITALIC)
    style.font.italic = False
    style.font.bold = False

    style = doc.styles["Strong"]
    func(style, FontFamily.BOLD)
    style.font.italic = False
    style.font.bold = True

    for i in range(9):
        style = doc.styles[f"Heading {i+1}"]
        func(style, FontFamily.ROMAN)
        style.font.color.rgb = COLORS["black"]

    style = doc.styles["Title"]
    func(style, FontFamily.ROMAN)
    style.font.color.rgb = COLORS["black"]

    return doc


@memoize
def load_data(csv_file: str) -> pd.DataFrame:
    return (
        Pipe(csv_file)
        .map(curry(pd.read_csv)(sep="\t")(names=HEADINGS.values()))
        .map(
            lambda x: x.sort_values(
                by=[HEADINGS[ColumnType.NAME], HEADINGS[ColumnType.DATE]]
            ),
        )
        .flush()
    )
