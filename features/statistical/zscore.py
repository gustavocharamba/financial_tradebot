import pandas as pd
import numpy as np

def getZscore(data: pd.Series, window):

    sma = data.rolling(window=window).mean()

    deviation = data.rolling(window=window).std()

    zscore = (data - sma) / deviation

    return zscore