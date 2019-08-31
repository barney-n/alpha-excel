import pandas as pd
import numpy as np
from alpha_vantage.timeseries import TimeSeries
import xlwings as xw
import datetime as dt

@xw.func
@xw.ret(index=True, header=False, expand='table')
def BDH(ticker, start_date, end_date):
	return daily_close(ticker, start_date, end_date)


def daily_close(ticker, start_date, end_date):
	
	av = TimeSeries(key='SJA94IVRELN62B8S', output_format='pandas')
	df = av.get_daily(symbol=ticker, outputsize='full')[0]

	df.index = [dt.datetime.strptime(i, '%Y-%m-%d').date() for i in df.index]
	
	start_date = fix_type(start_date)
	end_date = fix_type(end_date)

	while start_date not in df.index:
		start_date = start_date + dt.timedelta(days=1)
	while end_date not in df.index:
		end_date = end_date + dt.timedelta(days=1)

	return df[start_date:end_date]['4. close']


def string_to_date(date):
	date = dt.datetime.strptime(date, '%Y-%m-%d').date()
	return date

def fix_type(d):
	if d==str(d):
		return string_to_date(d)
	elif 'datetime.datetime' in str(type(d)):
		return d.date()

