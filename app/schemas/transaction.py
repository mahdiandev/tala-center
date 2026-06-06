from pydantic import BaseModel, ConfigDict
from utils.shamsi import Shamsi


class Transaction(BaseModel):
    """
    Pydantic schema representing a normalized bank deposit transaction.
    """
    amount: int
    date: Shamsi
    bank_name: str
    gold_price: int = 0

    model_config = ConfigDict(arbitrary_types_allowed=True)