import pandas as pd


def flag_rows(df: pd.DataFrame):
    cost_price = df[['costprice']]
    cost_outliers = cost_price[cost_price['costprice'] > cost_price.quantile(0.99)[0]]

    return df.loc[cost_outliers.index]

def remove_bars(df: pd.DataFrame):
    return df[df['srvcrsname'] != 'Bars']