from docx.document import Document
from pymonad.io import IO  # type:ignore

from templates import templates

if __name__ == "__main__":
    CSV_FILE = "报销表格.csv"
    DOC_FILE = "报销.docx"

    doc_file = (
        IO(
            lambda: input(
                f"Enter .docx file name (no suffix, press enter for {DOC_FILE}): "
            )
        )
        .map(lambda x: DOC_FILE if not x else x)
        .run()
    )
    csv_file = (
        IO(
            lambda: input(
                f"Enter .csv file name (no suffix, press enter for {CSV_FILE}): "
            )
        )
        .map(lambda x: CSV_FILE if not x else x)
        .run()
    )
    template = (
        IO(lambda: input(f"Select template (options: {list(templates.keys())}): "))
        .map(lambda x: templates[x])
        .run()
    )

    doc: Document = template(csv_file)
    doc.save(doc_file)
    print(f"Saved to {doc_file}.")
