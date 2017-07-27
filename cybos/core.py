import time

import psutil
import pyautogui
import pywinauto
from win32com import client


class CybosClient:

    __client__ = None

    def get_header(self, index):
        return self.__client__.GetHeaderValue(index)

    def get_data(self, row, column):
        return self.__client__.GetDataValue(row, column)

    def run(self):
        self.__client__.BlockRequest()

    def set_input_value(self, key, value):
        self.__client__.SetInputValue(key, value)


class StockChart(CybosClient):

    __client__ = client.Dispatch("CpSysDib.StockChart")


class StockTrader(CybosClient):

    __client__ = client.Dispatch("CpTrade.CpTd0311")


class StockUtil(CybosClient):

    __client__ = client.Dispatch("CpTrade.CpTdUtil")

    def __init__(self):
        self.__client__.TradeInit()


class Cybos:

    __stock_chart__ = None
    __stock_trader__ = None
    __stock_utill__ = None
    __bank_account_number__ = ''

    @property
    def stock_chart(self):
        if self.__stock_chart__ is None:
            self.__stock_chart__ = StockChart()
        return self.__stock_chart__

    @property
    def stock_trader(self):
        if self.__stock_trader__ is None:
            self.__stock_trader__ = StockTrader()
        assert self.__stock_utill__ is not None
        return self.__stock_trader__

    @property
    def stock_util(self):
        if self.__stock_utill__ is None:
            self.__stock_utill__ = StockUtil()
        return self.__stock_utill__

    @staticmethod
    def run_process(account_password, certification_password):
        app = pywinauto.Application()
        app.start('D:\DAISHIN\STARTER\\ncStarter.exe /prj:cp')
        time.sleep(1)
        pyautogui.typewrite('\n', interval=0.1)

        dialog = pywinauto.timings.WaitUntilPasses(20, 0.5, lambda: app.window(title='CYBOS Starter'))

        account_password_input = dialog.Edit2
        account_password_input.SetFocus()
        account_password_input.TypeKeys(account_password)

        certification_password_input = dialog.Edit3
        certification_password_input.SetFocus()
        certification_password_input.TypeKeys(certification_password)

        dialog.Button.Click()

        time.sleep(5)
        pyautogui.typewrite('\n', interval=0.5)
        time.sleep(10)

    def get_chart(self, code, count=10):
        self.stock_chart.set_input_value(0, code)
        self.stock_chart.set_input_value(1, ord('2'))
        # self.stock_chart.set_input_value(2, 'YYYYMMDD')
        # self.stock_chart.set_input_value(3, 'YYYYMMDD')
        self.stock_chart.set_input_value(4, count)
        self.stock_chart.set_input_value(5, (0, 2, 3, 4, 5, 8))
        self.stock_chart.set_input_value(6, ord('D'))
        self.stock_chart.set_input_value(9, ord('1'))

        self.stock_chart.run()
        rows = range(self.stock_chart.get_header(3))
        columns = range(self.stock_chart.get_header(1))

        results = []
        for row in rows:
            data = []
            for column in columns:
                data.append(self.stock_chart.get_data(column, row))
            results.append(data)
        return results

    def trade(self, trade_type: int, code: str, quantity: int, price: int, bank_account_number: str):
        self.stock_trader.set_input_value(0, trade_type)
        if bank_account_number == '':
            bank_account_number = self.__bank_account_number__
        self.stock_trader.set_input_value(1, bank_account_number)
        self.stock_trader.set_input_value(3, code)
        self.stock_trader.set_input_value(4, quantity)
        self.stock_trader.set_input_value(5, price)
        self.stock_trader.run()

    def sell(self, code, quantity, price, bank_account_number=''):
        self.trade(1, code, quantity, price, bank_account_number)

    def buy(self, code, quantity, price, bank_account_number=''):
        self.trade(2, code, quantity, price, bank_account_number)

    def __init__(self, account_password, certification_password, bank_account_number=''):
        if 'CpStart.exe' not in [p.name() for p in psutil.process_iter()]:
            self.run_process(account_password, certification_password)

        if bank_account_number == '':
            self.__bank_account_number__ = self.stock_util.__client__.AccountNumber[0]
        else:
            self.__bank_account_number__ = bank_account_number
