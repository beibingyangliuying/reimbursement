from enum import Enum, auto

import pandas as pd
from cytoolz.curried import curry, memoize  # type:ignore
from pymonad.reader import Pipe  # type:ignore


class Column(Enum):
    DATE = auto()
    NAME = auto()
    CATEGORY = auto()
    HASINVOICE = auto()
    DESCRIPTION = auto()
    AMOUNT = auto()


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
