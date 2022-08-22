import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import pandas as pd
import numpy as np

def function_1():
    df = pd.DataFrame(pd.read_excel(r'D:\Desktop\user_function_1.xlsx',engine='openpyxl'))
    # total_df = pd.DataFrame(pd.read_excel('D:\\test\\reslut_real.xlsx'))
    print(df)

function_1()