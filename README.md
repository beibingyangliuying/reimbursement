# Reimbursement

## Introduction

This project is designed to generate reimbursement document from reimbursement table automatically.

## How to use?

First, set up your `Python` virtual environment using [Poetry](https://python-poetry.org/). Next, prepare your `.csv` file in the following format:

| Date | Name | Category | Has Invoice | Description | Amount |
| --- | --- | --- | --- | --- | --- |
| 2024-8-31 | Your Name | Work | Yes | Your Description | 100.00 |

The delimiters of your `.csv` file should be tab `\t` and headers are not required (you could prepare a `.xlsx` file and paste it into the `.csv` file).

Finally, run `main.py` and enter the name of the `.docx` file, the name of the `.csv` file, and the name of the template.

## How to extend?

See `qiu.py` under `templates` package for intuition.

You should define a series of functions, each adding some contents to the document. These functions take up to two arguments: `doc` and `csv_file`. Using `curry` to decorate them, so that your document will be processed in a pipeline.

Next define the `main` function in the following format:

```python
def main(csv_file: str) -> Document:
    return (
        Pipe(init_blank_document())
        .map(beginning)
        .map(summary(csv_file=csv_file))
        .map(expense_statement(csv_file=csv_file))
        .map(confirm(csv_file=csv_file))
        .flush()
    )
```

Finally, add your `main` function in the `__init__.py` file to support templates choosing:

```python
from .qiu import main as qiu

templates = {"qiu": qiu}
__all__ = ["templates"]
```
