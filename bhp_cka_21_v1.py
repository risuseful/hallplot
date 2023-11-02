# -*- coding: utf-8 -*-
"""
Created on Tue Oct  3 11:00:01 2023

@author: akmalaulia
"""

import pandas as pd
from sklearn.preprocessing import StandardScaler
import numpy as np
from sklearn.neural_network import MLPRegressor
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from scipy.stats.stats import pearsonr
from sklearn.metrics import mean_squared_error
from math import sqrt
import pyodbc

# set global variables
global nnet, scaler_in, scaler_out

# set constant
min_date = '2023-04-01' # format: yyyy-mm-dd




# -----------------------------------------------------
#
#
# T R A I N I N G
#
#
# -----------------------------------------------------

# read data
d = pd.read_csv('train_cka21.csv')
nr = len(d) # number of rows

# set input and output
X = d[['Q_STBD', 'THP_PSIG']].values # input
y = d['BHP_PSI'].values # input

# feature scaling
scaler_in = StandardScaler()
scaler_in.fit(X)
X = scaler_in.transform(X)

# output scaling
scaler_out = StandardScaler()
y = y.reshape(-1,1) # need to reshape in order to use scaler_out
scaler_out.fit(y)
y = scaler_out.transform(y)

# construct neural net
nnet = MLPRegressor(solver='lbfgs', alpha=1e-5, hidden_layer_sizes=(7, 2), random_state=1)

# train neural net
nnet.fit(X,y) #






# -----------------------------------------------------
#
#
# P R E D I C T I O N
#
#
# -----------------------------------------------------

# note: 
#       - input file: pulled from database.  Columns -> Date, Q_STBD, THP_PSIA
#       - output file: out_pred_cka21.csv. With columns -> Date, Q_STBD, THP_PSIA, BHP_PSIA, cum_Q_STB, cum_BHP_PSIA

# construct SQL query
# query = "SELECT cast([TIMESTAMP] as Date) as Date ,[VOL_INJ_CK21] as Q_STBD ,([INJ_BP_CK21]*14.5+14.7) as THP_PSIA " # in PSIA
query = "SELECT cast([TIMESTAMP] as Date) as Date , "
query = query + "coalesce([VOL_INJ_CK21], avg([VOL_INJ_CK21]) over ()) as Q_STBD , "
query = query + "coalesce(([INJ_BP_CK21]*14.5), avg([INJ_BP_CK21]*14.5) over ()) as THP_PSIG " # in PSIG
query = query + "FROM [*****].[*****].[*****] where TIMESTAMP > " + "'" + min_date + "'" + " order by TIMESTAMP asc"

# use SQL query to pull from database
conn = pyodbc.connect(driver='*****', server='*****', user='*****', password='*****', database='*****')
cursor = conn.cursor()
d_pred = pd.read_sql_query(query, conn)

# set input matrix for nnet-based prediction
X_pred = d_pred[['Q_STBD', 'THP_PSIG']].values # input
X_pred = scaler_in.transform(X_pred) # use the previous scaler for X

# predict BHP from X_pred
y_pred = nnet.predict(X_pred) # predict

# inverse transform
y_pred = y_pred.reshape(-1,1) # reshape predicted
y_pred = scaler_out.inverse_transform(y_pred)


# compile out_pred_cka21.csv
res = pd.DataFrame()
res['Date'] = d_pred['Date']
res['Q_STBD'] = d_pred['Q_STBD']
res['THP_PSIG'] = d_pred['THP_PSIG']
res['BHP_PSI'] = y_pred.reshape(-1)
res['cum_Q_STBD'] = res['Q_STBD'].cumsum()
res['cum_BHP_PSI'] = res['BHP_PSI'].cumsum()
res.to_csv('out_pred_cka21.csv')
