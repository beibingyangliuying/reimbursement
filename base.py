from enum import Enum, auto

import pandas as pd
from cytoolz.curried import curry, memoize  # type:ignore
from docx.document import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from pymonad.reader import Pipe  # type:ignore


class Column(Enum):
    DATE = auto()
    NAME = auto()
    CATEGORY = auto()
    HASINVOICE = auto()
    DESCRIPTION = auto()
    AMOUNT = auto()


class Color(Enum):
    BLACK = auto()
    RED = auto()
    BLUE = auto()
    GREEN = auto()


def color(color_type: Color) -> RGBColor:
    match color_type:
        case Color.BLACK:
            return RGBColor(0, 0, 0)
        case Color.RED:
            return RGBColor(255, 0, 0)
        case Color.BLUE:
            return RGBColor(0, 0, 255)
        case Color.GREEN:
            return RGBColor(0, 255, 0)


class FontSize(Enum):
    初号 = auto()
    小初 = auto()
    一号 = auto()
    小一 = auto()
    二号 = auto()
    小二 = auto()
    三号 = auto()
    小三 = auto()
    四号 = auto()
    小四 = auto()
    五号 = auto()
    小五 = auto()
    六号 = auto()
    小六 = auto()
    七号 = auto()
    八号 = auto()


def font_size(font_type: FontSize) -> Pt:
    match font_type:
        case FontSize.初号:
            return Pt(42)
        case FontSize.小初:
            return Pt(36)
        case FontSize.一号:
            return Pt(26)
        case FontSize.小一:
            return Pt(24)
        case FontSize.二号:
            return Pt(22)
        case FontSize.小二:
            return Pt(18)
        case FontSize.三号:
            return Pt(16)
        case FontSize.小三:
            return Pt(15)
        case FontSize.四号:
            return Pt(14)
        case FontSize.小四:
            return Pt(12)
        case FontSize.五号:
            return Pt(10.5)
        case FontSize.小五:
            return Pt(9)
        case FontSize.六号:
            return Pt(7.5)
        case FontSize.小六:
            return Pt(6.5)
        case FontSize.七号:
            return Pt(5.5)
        case FontSize.八号:
            return Pt(5)


class FontFamily(Enum):
    ROMAN = auto()
    ITALIC = auto()
    BOLD = auto()


def font_family(font_type: FontFamily) -> tuple[str, str]:
    match font_type:
        case FontFamily.ROMAN:
            return ("Times New Roman", "宋体")
        case FontFamily.ITALIC:
            return ("Times New Roman", "楷体")
        case FontFamily.BOLD:
            return ("Arial", "黑体")


def set_style_font(style, font_type: FontFamily) -> bool:
    (western, asian) = font_family(font_type)
    try:
        style.font.name = western
        style.font._element.rPr.rFonts.set(qn("w:eastAsia"), asian)
    except AttributeError:
        return False
    return True


def init_blank_document() -> Document:
    from docx import Document as create_document

    doc = create_document()
    for style in doc.styles:
        set_style_font(style, FontFamily.ROMAN)

    style = doc.styles["Emphasis"]
    set_style_font(style, FontFamily.ITALIC)
    style.font.italic = False
    style.font.bold = False

    style = doc.styles["Strong"]
    set_style_font(style, FontFamily.BOLD)
    style.font.italic = False
    style.font.bold = True

    return doc


@memoize
def load_data(csv_file: str) -> pd.DataFrame:
    return (
        Pipe(csv_file)
        .map(
            curry(pd.read_csv)(sep="\t")(
                names=[
                    Column.DATE,
                    Column.NAME,
                    Column.CATEGORY,
                    Column.HASINVOICE,
                    Column.DESCRIPTION,
                    Column.AMOUNT,
                ]
            )
        )
        .map(
            lambda x: x.sort_values(
                by=[
                    Column.NAME,
                    Column.DATE,
                    Column.CATEGORY,
                ]
            ),
        )
        .flush()
    )
