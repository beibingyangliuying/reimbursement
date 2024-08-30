import datetime
from itertools import product

from cytoolz.curried import curry  # type:ignore
from docx.document import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pymonad.reader import Pipe  # type:ignore

from .base import COLORS, FONTSIZE, HEADINGS, ColumnType, init_blank_document, load_data


@curry
def expense_statement(doc: Document, csv_file: str) -> Document:
    def func(x: int) -> ColumnType:
        match x:
            case 0:
                return ColumnType.NAME
            case 1:
                return ColumnType.DATE
            case 2:
                return ColumnType.DESCRIPTION
            case 3:
                return ColumnType.AMOUNT
            case _:
                raise ValueError(f"Invalid index: {x}.")

    doc.add_heading(text="费用明细", level=1)

    headings = ("姓名", "日期", "事项", "费用", "合计")
    for i, (category, df) in enumerate(
        load_data(csv_file).groupby(HEADINGS[ColumnType.CATEGORY])
    ):
        doc.add_paragraph(f"（{i + 1}）{category}费用：")
        paragraph_summary = doc.add_paragraph()
        paragraph_summary.add_run("支付组成：").bold = True

        heading_rows = 1
        rows = df.shape[0] + heading_rows
        table = doc.add_table(rows=rows, cols=len(headings), style="Table Grid")
        # Fill headings
        for i, heading in enumerate(headings):
            table.cell(0, i).text = heading
        # Fill isolate data
        for i, j in product(range(heading_rows, rows), range(1, 4)):
            table.cell(i, j).text = str(df[HEADINGS[func(j)]].iloc[i - heading_rows])
        # todo: Fill summary data
        i = heading_rows
        for name, df in df.groupby(HEADINGS[ColumnType.NAME]):
            cell_name = table.cell(i, 0)
            cell_name.text = name

            cell_sum = table.cell(i, len(headings) - 1)
            amount_sum = round(df[HEADINGS[ColumnType.AMOUNT]].sum(), 2)
            cell_sum.text = str(amount_sum)

            i += df.shape[0]
            cell_name.merge(table.cell(i - 1, 0))
            cell_sum.merge(table.cell(i - 1, len(headings) - 1))

            run = paragraph_summary.add_run(f"{name} ")
            run.bold = True
            run = paragraph_summary.add_run(f"￥{amount_sum} ")
            run.bold = True
            run.font.color.rgb = COLORS["blue"]

        doc.add_paragraph()

    return doc


@curry
def summary(doc: Document, csv_file: str) -> Document:
    names = list(load_data(csv_file).groupby(HEADINGS[ColumnType.NAME]).groups.keys())

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(f"以上支付明细详单\r经出差{len(names)}人确认无误！")
    run.style = "Strong"  # type:ignore
    run.font.size = FONTSIZE["一号"]
    run.font.color.rgb = COLORS["red"]

    doc.add_paragraph()

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run(str(" ".join(names)))
    run.style = "Emphasis"  # type:ignore
    run.bold = True

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run(f"{datetime.datetime.now().strftime('%Y年%m月%d日')}")
    run.font.bold = True

    return doc


@curry
def beginning(doc: Document, title: str) -> Document:
    doc.add_paragraph(title, style="Title")
    paragraph = doc.add_paragraph()
    paragraph.add_run("出差时间：").bold = True

    paragraph = doc.add_paragraph()
    paragraph.add_run("出差地点：").bold = True

    paragraph = doc.add_paragraph()
    paragraph.add_run("报销时间：").bold = True

    doc.add_paragraph("-" * 50)
    doc.add_paragraph("出差前预支款：")
    doc.add_paragraph("-" * 50)

    return doc


def main(csv_file: str, **kwargs) -> Document:  # type:ignore
    title = kwargs.get("title", "填写标题")

    return (
        Pipe(init_blank_document())
        .map(beginning(title=title))
        .map(expense_statement(csv_file=csv_file))
        .map(summary(csv_file=csv_file))
        .flush()
    )
