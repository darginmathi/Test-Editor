import pandas as pd
from .model import TableModel

class ObjRepoModel(TableModel):

    TYPE_COL = 0
    NAME_COL = 1
    XPATH_COL = 2
    LOCATOR_COL = 3

    def __init__(self, df: pd.DataFrame = None):
        super().__init__(df=df)

    @classmethod
    def create_preset(cls):
        columns = [str(i) for i in range(4)]
        data = [
            ["Type", "User friendly name of Object", "By-Type", "Webdriver friendly name of Object"],
            ["Link", "lnkAdmin", "XPATH", "//*[@id=\"page-admin\"]"],
            ["END", "", "", ""]
        ]
        return pd.DataFrame(data, columns=columns)
