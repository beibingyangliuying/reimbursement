from docx.document import Document
from pymonad.io import IO  # type:ignore

from templates import qiu

if __name__ == "__main__":
    CSV_FILE = "报销表格.csv"
    DOC_FILE = "报销.docx"

    doc_file = (
        IO(
            lambda: input(
                "Enter .docx file name (no suffix, press enter for default): "
            )
        )
        .map(lambda x: DOC_FILE if not x else x)
        .run()
    )
    csv_file = (
        IO(lambda: input("Enter .csv file name (no suffix, press enter for default): "))
        .map(lambda x: CSV_FILE if not x else x)
        .run()
    )

    # todo: Select template.
    doc: Document = qiu(csv_file)
    doc.save(doc_file)
    print(f"Saved to {doc_file}.")
