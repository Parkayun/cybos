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

    __client__ = client.Dispatch("CpTrade.CpTdUtil")

    def __init__(self):
        self.__client__.TradeInit()


class Cybos:

    __stock_chart__ = None
    __stock_trader__ = None

    @property
    def stock_chart(self):
        if self.__stock_chart__ is None:
            self.__stock_chart__ = StockChart()
        return self.__stock_chart__

    @property
    def stock_trader(self):
        if self.__stock_trader__ is None:
            self.__stock_trader__ = StockTrader()
        return self.__stock_trader__

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