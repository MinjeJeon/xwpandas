import os
import csv
import xlwings as xw
import numpy as np
import pandas as pd

from pathlib import Path
from typing import Union, Iterable, Optional
from pandas import DataFrame, Series
from xlwings import Book
from xlwings.constants import Constants as C

from .utils import is_iterable, safe_path, temp_path


class Xwhandler:

    def __init__(self, path: Optional[Union[str, Path]], mode: str='r', use_existing_app: bool = True,
                 close_on_exit: bool= True):
        if mode in ['r', 'w']:
            self.mode = mode
        else:
            raise ValueError()

        if path is None:
            if mode == 'w':
                self.path = None
                self.close_on_exit = False
            elif mode == 'r':
                raise ValueError
        else:
            self.path = str(path)
            self.close_on_exit = close_on_exit

        self.use_existing_app = use_existing_app
        self.wb = None
        self.app = None
        self.closed = False

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()

    def save(self, safe: bool=False):
        if safe:
            self.wb.save(safe_path(self.path))
        else:
            self.wb.save(self.path)

    def open(self):
        if xw.apps and self.use_existing_app:
            self.app = xw.apps.active
            if self.mode == 'r':
                self.wb = self.app.books.open(self.path)
            elif self.mode == 'w':
                self.wb = self.app.books.add()
        else:
            if self.mode == 'r':
                self.wb = xw.Book(self.path)
            elif self.mode == 'w':
                self.wb = xw.Book()
            self.app = self.wb.app
        self.app.screen_updating = False

    def close(self):
        self.app.screen_updating = True
        if self.close_on_exit:
            self.closed = True
            self.wb.close()
            if not self.app.books:
                self.app.quit()


def autofit_sheet(sheet: Optional[xw.Sheet]=None, max_col_width: Optional[float] = 90.0):
    if sheet is None:
        last_cell_address = xw.sheets.active.cells.api.SpecialCells(C.xlLastCell).Address
        whole_table = xw.sheets.active.range('A1', last_cell_address)
    else:
        last_cell_address = sheet.cells.api.SpecialCells(C.xlLastCell).Address
        whole_table = sheet.range('A1', last_cell_address)
    whole_table.columns.autofit()
    for col in whole_table.columns:
        if max_col_width is not None and col.column_width > max_col_width:
            col.column_width = max_col_width


def read(path: Union[str, Path], sheets: Union[int, str, Iterable[Union[int, str]]] = 0, use_existing_app: bool = True,
         close_on_exit: bool= True) -> Union[DataFrame, dict]:

    res = {}
    with Xwhandler(path, mode='r', use_existing_app=use_existing_app, close_on_exit=close_on_exit) as w:
        if sheets is None:
            sheets_iterable = [x.name for x in w.wb.sheets]
        elif is_iterable(sheets):
            sheets_iterable = sheets
        else:
            sheets_iterable = [sheets]
        for sheet in sheets_iterable:
            xl_range = w.wb.sheets[sheet].range('A1').expand().options(DataFrame)
            res[sheet] = xl_range.value
    if len(res) == 1:
        for df in res.values():
            return df
    else:
        return res


def check_dataframe(df:Union[DataFrame, Series]) -> DataFrame:
    if df.empty:
        raise ValueError('DataFrame is Empty')
    df = df if isinstance(df, DataFrame) else df.to_frame()
    return df


def excel_header(df:pd.DataFrame, sheet:xw.Sheet) -> xw.Range:
    try:
        current_app = sheet.book.app
        temp_excel_path = temp_path('.xlsx')
        df.head(1).to_excel(temp_excel_path)
        xl_header = current_app.books.open(temp_excel_path).sheets[0]
        xl_header_last_cell_address = xl_header.cells.api.SpecialCells(C.xlLastCell).Address
        xl_header_last_row = xw.Range(xl_header_last_cell_address).row
        xl_header.range('A1', xl_header_last_cell_address).api.Copy(
            sheet.range('A1').api
        )
        sheet.activate()
        sheet.api.Rows(xl_header_last_row).Delete()
        return sheet.cells(xl_header_last_row, 1)
    except Exception as e:
        raise e
    finally:
        xl_header.book.close()
        os.unlink(temp_excel_path)


def df_to_csv(df: DataFrame):
    no_index_df = df.reset_index()
    temp_csv_path = temp_path('.csv')
    no_index_df_dtypes = no_index_df.dtypes
    no_index_df.to_csv(temp_csv_path, header=False, index=False, encoding='utf-8', quoting=csv.QUOTE_NONNUMERIC)
    return temp_csv_path, no_index_df_dtypes


def _df_toxlwings(df: Union[DataFrame, Series], path: Optional[Union[str, Path]] = None, safe: bool=False,
                  max_col_width: Optional[float] = 90.0, autofit: bool = True,
                  use_existing_app: bool = True, close_on_exit: bool= False) -> Optional[Book]:

    df = check_dataframe(df)
    df_length = len(df)
    no_index_df = df.reset_index()

    with Xwhandler(path, mode='w', use_existing_app=use_existing_app, close_on_exit=close_on_exit) as w:
        activesheet = w.wb.sheets.active
        target_range = excel_header(df, activesheet)
        start_row = target_range.row
        for colnum_current, col_name_and_series in enumerate(no_index_df.iteritems(), 1):
            name, col_series = col_name_and_series
            col_series.name = None
            pos_start = (start_row, colnum_current)
            pos_end = (start_row + df_length - 1, colnum_current)
            if pos_end[0] > 1024768:
                raise ValueError('DataFrame exceeds row limit of excel')
            if col_series.dtype == np.object_:
                activesheet.range(pos_start, pos_end).number_format = '@'
            activesheet.range(pos_start).options(Series, index=False, header=False).value = col_series
        if autofit:
            autofit_sheet(activesheet, max_col_width=max_col_width)
        if path is not None:
            w.save(safe=safe)
        if not w.close_on_exit:
            wb = w.wb
            return wb


def _df_toxlwings_csv(df: Union[DataFrame, Series], path: Optional[Union[str, Path]] = None, safe: bool=False,
                      max_col_width: Optional[float] = 90.0, autofit: bool = True,
                      use_existing_app: bool = True, close_on_exit: bool= False) -> Optional[Book]:

    df = check_dataframe(df)
    with Xwhandler(path, mode='w', use_existing_app=use_existing_app, close_on_exit=close_on_exit) as w:
        activesheet = w.wb.sheets.active
        target_range = excel_header(df, activesheet)
        temp_csv_path, no_index_df_dtypes = df_to_csv(df)
        no_index_col_dtypes = np.where(no_index_df_dtypes == np.object_, 2, 1).tolist()

        # add querytable
        try:
            t = activesheet.api.QueryTables.Add(
                Connection='TEXT;{}'.format(temp_csv_path),
                Destination=activesheet.range(target_range.address).api
            )
            t.Name = "querytable_from_xwpandas"
            t.FieldNames = False
            t.FillAdjacentFormulas = False
            t.PreserveFormatting = True
            t.AdjustColumnWidth = False
            t.RefreshOnFileOpen = False
            t.RefreshStyle = 0  # xlOverwriteCells
            t.SaveData = True
            t.TextFilePlatform = 65001
            t.TextFileStartRow = 1
            t.TextFileParseType = 1  # xlDelimited
            t.TextFileTextQualifier = 1  # xlTextQualifierDoubleQuote
            t.TextFileConsecutiveDelimiter = False
            t.TextFileTabDelimiter = False
            t.TextFileSemicolonDelimiter = False
            t.TextFileCommaDelimiter = True
            t.TextFileSpaceDelimiter = False
            t.TextFileColumnDataTypes = no_index_col_dtypes
            t.TextFileTrailingMinusNumbers = True
            t.Refresh(BackgroundQuery=False)
            if autofit:
                autofit_sheet(activesheet, max_col_width=max_col_width)
            if path is not None:
                w.save(safe=safe)
            if not w.close_on_exit:
                wb = w.wb
                return wb
        except Exception as e:
            raise e
        finally:
            os.unlink(temp_csv_path)


def save(df: DataFrame, path: Union[str, Path], method='xlwings', max_col_width: Optional[float] = 90.0,
         autofit: bool = True, use_existing_app: bool = True, close_on_exit: bool= True) -> Optional[Book]:
    if method == 'xlwings':
        return _df_toxlwings(df=df, path=path, max_col_width=max_col_width, autofit=autofit,
                             use_existing_app=use_existing_app, close_on_exit=close_on_exit)
    elif method == 'csv':
        return _df_toxlwings_csv(df=df, path=path, max_col_width=max_col_width, autofit=autofit,
                                 use_existing_app=use_existing_app, close_on_exit=close_on_exit)
    else:
        raise ValueError("method only supports xlwings or csv")


def view(df: DataFrame, method='xlwings', max_col_width: Optional[float] = 90.0, autofit: bool = True,
         use_existing_app: bool = True) -> Optional[Book]:
    if method == 'xlwings':
        return _df_toxlwings(df=df, max_col_width=max_col_width, autofit=autofit, use_existing_app=use_existing_app)
    elif method == 'csv':
        return _df_toxlwings_csv(df=df, max_col_width=max_col_width, autofit=autofit, use_existing_app=use_existing_app)
    else:
        raise ValueError("method only supports xlwings or csv")

