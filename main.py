from coinMarket_research import *
from openpyxl import Workbook
from openpyxl import load_workbook
import os
import time
import sys

coin_researcher = CoinMarketResearch()
coin_researcher.write_market_research()
