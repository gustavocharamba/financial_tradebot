import numpy as np
import pandas as pd
from statsmodels.api import OLS, add_constant
from statsmodels.tsa.stattools import adfuller

def getCointegration(data_x: pd.Series, data_y: pd.Series, log_prices: bool = True):
    # Remove NaNs
    data = pd.concat([data_y, data_x], axis=1).dropna()
    y, x = data.iloc[:, 0], data.iloc[:, 1]

    if log_prices:
        y = np.log(y)
        x = np.log(x)

    # Linear Regression y = a + bx
    X = add_constant(x)
    model = OLS(y, X).fit()
    alpha, beta = model.params

    # Regression Residual
    resid = y - (alpha + beta * x)

    # ADF Test
    try:
        adf_stat, p_value, *_ = adfuller(resid)
    except Exception:
        adf_stat, p_value = np.nan, np.nan

    return {
        'Alpha': alpha,
        'Beta': beta,
        'ADF_Stat': adf_stat,
        'P_Value': p_value,
        'Is_Cointegrated': (p_value < 0.05 if not np.isnan(p_value) else False),
        'Residuals': resid
    }
