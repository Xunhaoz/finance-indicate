import talib
import pandas as pd
from pandas import ExcelWriter
import numpy as np
import copy


def cal_KDJ(df):
    rsv = (df['close'] - df['low']) / (df['high'] - df['low']) * 100
    k = [50, ]
    d = [50, ]
    for key, value in enumerate(rsv):
        if key == 0:
            continue

        k.append((k[key - 1] * 2 + value) / 3)
        d.append((d[key - 1] * 2 + k[key]) / 3)

    k = np.array(k)
    d = np.array(d)
    j = 3 * k - 2 * d
    return rsv, k, d, j


def cal_RSI(df):
    df = df['close'] - df['close'].shift(1)
    up = pd.Series(index=df.index, data=df[df > 0])
    up = up.fillna(0)
    down = pd.Series(index=df.index, data=-df[df < 0])
    down = down.fillna(0)

    up_mean = []
    down_mean = []
    for i in range(10, len(df) + 1):
        up_mean.append(np.mean(up.values[i - 9:i]))
        down_mean.append(np.mean(down.values[i - 9:i]))

    rsi = [None for i in range(9)]
    for i in range(len(up_mean)):
        rsi.append(100 * up_mean[i] / (up_mean[i] + down_mean[i]))
    rsi = np.array(rsi)

    return rsi


def cal_MACD(df):
    di = (df['high'] + df['low'] + df['close'] * 2) / 4
    EMA12 = []
    EMA26 = []
    DIF = []
    MACD = []
    OSC = []
    for key, value in enumerate(di):
        if key < 11:
            EMA12.append(None)
        elif key == 11:
            EMA12.append(di[0:12].mean())
        else:
            EMA12.append(EMA12[len(EMA12) - 1] * 11 / 13 + value * 2 / 13)
        if key < 25:
            EMA26.append(None)
        elif key == 25:
            EMA26.append(di[0:26].mean())
        else:
            EMA26.append(EMA26[len(EMA26) - 1] * 25 / 27 + value * 2 / 27)
        if EMA12[len(EMA12) - 1] and EMA26[len(EMA26) - 1]:
            DIF.append(EMA12[len(EMA12) - 1] - EMA26[len(EMA26) - 1])
        else:
            DIF.append(None)
        if key < 33:
            MACD.append(None)
        elif key == 33:
            MACD.append(np.array(DIF[key - 8:key + 1]).mean())
        else:
            MACD.append(MACD[len(MACD) - 1] * 8 / 10 + DIF[len(DIF) - 1] * 2 / 10)
        if MACD[len(MACD) - 1] and DIF[len(DIF) - 1]:
            OSC.append(DIF[len(DIF) - 1] - MACD[len(MACD) - 1])
        else:
            OSC.append(None)

    return DIF, MACD, OSC


def cal_bband(df):
    df = df['close']
    real_SMA = talib.SMA(df, timeperiod=20)
    sigma = [None for i in range(19)]
    for i in range(20, len(df) + 1):
        sigma.append(df[i - 20: i].std())

    lowerband = []
    upperband = []

    for key, value in enumerate(real_SMA):
        if value and sigma[key]:
            upperband.append(value + 2 * sigma[key])
            lowerband.append(value - 2 * sigma[key])
        else:
            upperband.append(None)
            lowerband.append(None)

    return upperband, real_SMA, lowerband


def cal_HMA(df):
    real_WMA = talib.WMA(df['close'], timeperiod=5)
    WMA_1 = [None for _ in range(10)]
    HMA = [None for _ in range(10)]
    for key in range(10, len(real_WMA)):
        WMA_1.append(real_WMA[key - 3: key].mean() * 2 - real_WMA[key - 6: key].mean())
        HMA.append((WMA_1[len(WMA_1) - 1] + real_WMA[key - 3: key].mean()) / 2)
    return HMA


if __name__ == "__main__":
    df = pd.read_csv('source.csv')

    # word topic 1
    real_SMA = talib.SMA(df['close'], timeperiod=5)
    real_EMA = talib.EMA(df['close'], timeperiod=5)
    real_WMA = talib.WMA(df['close'], timeperiod=5)

    # word topic 2
    # k, d, j count by ta-lib / cannot meet customers expectation
    slowk, slowd = talib.STOCH(df['high'], df['low'], df['close'],
                               fastk_period=1, slowk_period=2, slowk_matype=0, slowd_period=2, slowd_matype=0)
    slowj = 3 * slowk - 2 * slowd
    # k, d, j count by function / works successfully
    fun_rsv, fun_k, fun_d, fun_j = cal_KDJ(df)

    # word topic 3
    # RSI / cannot meet customers expectation
    real_RSI = talib.RSI(df['close'], timeperiod=9)
    # RSI / works successfully
    fun_rsi = cal_RSI(df)

    # word topic 4
    # MACD / cannot meet customers expectation
    macd, macdsignal, macdhist = talib.MACD(df['close'], signalperiod=9)
    # MACD / works successfully
    fun_dif, fun_macd, fun_osc = cal_MACD(df)

    # word topic 5
    # BBANDS / cannot meet customers expectation
    upperband, middleband, lowerband = talib.BBANDS(df['close'], timeperiod=20, nbdevup=2, nbdevdn=2, matype=0)
    # BBAND / works successfully
    fun_upperband, fun_middleband, fun_lowerband = cal_bband(df)

    # word topic 6
    fun_HMA = cal_HMA(df)

    df_MA = df.copy(deep=True)
    df_KD_Ta = df.copy(deep=True)
    df_KD_fun = df.copy(deep=True)
    df_RSI_Ta = df.copy(deep=True)
    df_RSI_fun = df.copy(deep=True)
    df_MACD_Ta = df.copy(deep=True)
    df_MACD_fun = df.copy(deep=True)
    df_BBANDS_Ta = df.copy(deep=True)
    df_BBANDS_fun = df.copy(deep=True)

    df_MA.insert(len(df_MA.columns), column='SMA', value=real_SMA)
    df_MA.insert(len(df_MA.columns), column='EMA', value=real_EMA)
    df_MA.insert(len(df_MA.columns), column='WMA', value=real_WMA)
    df_MA.insert(len(df_MA.columns), column='HMA', value=fun_HMA)

    df_KD_Ta.insert(len(df_KD_Ta.columns), column='K', value=slowk)
    df_KD_Ta.insert(len(df_KD_Ta.columns), column='D', value=slowd)
    df_KD_Ta.insert(len(df_KD_Ta.columns), column='J', value=slowj)

    df_KD_fun.insert(len(df_KD_fun.columns), column='RSV', value=fun_rsv)
    df_KD_fun.insert(len(df_KD_fun.columns), column='K', value=fun_k)
    df_KD_fun.insert(len(df_KD_fun.columns), column='D', value=fun_d)
    df_KD_fun.insert(len(df_KD_fun.columns), column='J', value=fun_j)

    df_RSI_Ta.insert(len(df_RSI_Ta.columns), column='RSI', value=real_RSI)

    df_RSI_fun.insert(len(df_RSI_fun.columns), column='RSI', value=fun_rsi)

    df_MACD_Ta.insert(len(df_MACD_Ta.columns), column='DIF', value=macd)
    df_MACD_Ta.insert(len(df_MACD_Ta.columns), column='MACD', value=macdsignal)
    df_MACD_Ta.insert(len(df_MACD_Ta.columns), column='OSC', value=macdhist)

    df_MACD_fun.insert(len(df_MACD_fun.columns), column='DIF', value=fun_dif)
    df_MACD_fun.insert(len(df_MACD_fun.columns), column='MACD', value=fun_macd)
    df_MACD_fun.insert(len(df_MACD_fun.columns), column='OSC', value=fun_osc)

    df_BBANDS_Ta.insert(len(df_BBANDS_Ta.columns), column='upperband', value=upperband)
    df_BBANDS_Ta.insert(len(df_BBANDS_Ta.columns), column='middleband', value=middleband)
    df_BBANDS_Ta.insert(len(df_BBANDS_Ta.columns), column='lowerband', value=lowerband)

    df_BBANDS_fun.insert(len(df_BBANDS_fun.columns), column='upperband', value=fun_upperband)
    df_BBANDS_fun.insert(len(df_BBANDS_fun.columns), column='middleband', value=fun_middleband)
    df_BBANDS_fun.insert(len(df_BBANDS_fun.columns), column='lowerband', value=fun_lowerband)

    with ExcelWriter('product.xlsx') as writer:
        df_MA.to_excel(writer, sheet_name='MA')
        df_KD_Ta.to_excel(writer, sheet_name='KDJ_Ta')
        df_KD_fun.to_excel(writer, sheet_name='KDJ_fun')
        df_RSI_Ta.to_excel(writer, sheet_name='RSI_Ta')
        df_RSI_fun.to_excel(writer, sheet_name='RSI_fun')
        df_MACD_Ta.to_excel(writer, sheet_name='MACD')
        df_MACD_fun.to_excel(writer, sheet_name='MACD_fun')
        df_BBANDS_Ta.to_excel(writer, sheet_name='布林通道_Ta')
        df_BBANDS_fun.to_excel(writer, sheet_name='布林通道_fun')
        writer.save()
