# -*- coding: utf-8 -*-
import abc
import io
import tempfile

from pywinauto import win32defines
from typing import TYPE_CHECKING, Dict, List

import pandas as pd
import pywinauto.clipboard
from .utils import SetForegroundWindow, ShowWindow
import pywinauto
import logging
try:
    import StringIO
except:
    from io import StringIO

from .log import log
from .utils.captcha import captcha_recognize


if TYPE_CHECKING:
    # pylint: disable=unused-import
    from . import clienttrader


class IGridStrategy(abc.ABC):
    @abc.abstractmethod
    def get(self, control_id: int) -> List[Dict]:
        """
        获取 gird 数据并格式化返回

        :param control_id: grid 的 control id
        :return: grid 数据
        """
        pass


class BaseStrategy(IGridStrategy):
    def __init__(self, trader: "clienttrader.IClientTrader") -> None:
        self._trader = trader

    @abc.abstractmethod
    def get(self, control_id: int) -> List[Dict]:
        """
        :param control_id: grid 的 control id
        :return: grid 数据
        """
        pass

    def _get_grid(self, control_id: int):
        grid = self._trader.main.window(
            control_id=control_id, class_name="CVirtualGridCtrl"
        )
        return grid

    def _set_foreground(self, grid=None):
        if grid is None:
            grid = self._trader.main
        if grid.has_style(pywinauto.win32defines.WS_MINIMIZE):  # if minimized
            ShowWindow(grid.wrapper_object(), 9)  # restore window state
        else:
            SetForegroundWindow(grid.wrapper_object())  # bring to front


class Copy(BaseStrategy):
    """
    通过复制 grid 内容到剪切板z再读取来获取 grid 内容
    """
    _need_captcha_reg = True

    def get(self, control_id: int) -> List[Dict]:
        grid = self._get_grid(control_id)
        self._set_foreground(grid)
        grid.type_keys("^A^C", set_foreground=False)
        content = self._get_clipboard_data()
        return self._format_grid_data(content)

    def _format_grid_data(self, data: str) -> List[Dict]:
        try:
            df = pd.read_csv(
                io.StringIO(data),
                delimiter="\t",
                dtype=self._trader.config.GRID_DTYPE,
                na_filter=False,
            )
            return df.to_dict("records")
        except:
            Copy._need_captcha_reg = True

    def _get_clipboard_data(self) -> str:
        if Copy._need_captcha_reg:
            if self._trader.app.top_window().window(class_name='Static', title_re="验证码").exists(timeout=1):
                file_path = "tmp.png"
                count = 5
                found = False
                while count > 0:
                    self._trader.app.top_window().window(control_id=0x965, class_name='Static').\
                        capture_as_image().save(file_path)  # 保存验证码

                    captcha_num = captcha_recognize(file_path)  # 识别验证码
                    log.info("captcha result-->" + captcha_num)
                    if len(captcha_num) == 4:
                        self._trader.app.top_window().window(control_id=0x964, class_name='Edit').set_text(captcha_num)  # 模拟输入验证码

                        self._trader.app.top_window().set_focus()
                        pywinauto.keyboard.SendKeys("{ENTER}")   # 模拟发送enter，点击确定
                        try:
                            log.info(self._trader.app.top_window().window(control_id=0x966, class_name='Static').window_text())
                        except Exception as ex:       # 窗体消失
                            log.exception(ex)
                            found = True
                            break
                    count -= 1
                    self._trader.wait(0.1)
                    self._trader.app.top_window().window(control_id=0x965, class_name='Static').click()
                if not found:
                    self._trader.app.top_window().Button2.click()  # 点击取消
            else:
                Copy._need_captcha_reg = False
        count = 5
        while count > 0:
            try:
                return pywinauto.clipboard.GetData()
            # pylint: disable=broad-except
            except Exception as e:
                count -= 1
                log.exception("%s, retry ......", e)


class WMCopy(Copy):
    """
    通过复制 grid 内容到剪切板z再读取来获取 grid 内容
    """

    def get(self, control_id: int) -> List[Dict]:
        grid = self._get_grid(control_id)
        grid.post_message(win32defines.WM_COMMAND, 0xe122, 0)
        self._trader.wait(0.1)
        content = self._get_clipboard_data()
        return self._format_grid_data(content)



class Xls(BaseStrategy):
    """
    通过将 Grid 另存为 xls 文件再读取的方式获取 grid 内容，
    用于绕过一些客户端不允许复制的限制
    """

    def get(self, control_id: int) -> List[Dict]:
        grid = self._get_grid(control_id)

        # ctrl+s 保存 grid 内容为 xls 文件
        self._set_foreground(grid)  # setFocus buggy, instead of SetForegroundWindow
        grid.type_keys("^s", set_foreground=False)
        self._trader.wait(0.5)

        temp_path = tempfile.mktemp(suffix=".csv")
        self._set_foreground(self._trader.app.top_window())
        self._trader.app.top_window().type_keys(self.normalize_path(temp_path), set_foreground=False)

        # alt+s保存，alt+y替换已存在的文件
        self._set_foreground(self._trader.app.top_window())
        self._trader.app.top_window().type_keys("%{s}%{y}", set_foreground=False)
        # Wait until file save complete otherwise pandas can not find file
        self._trader.wait(0.2)
        return self._format_grid_data(temp_path)

    def normalize_path(self, temp_path: str) -> str:
        return temp_path.replace('~', '{~}')

    def _format_grid_data(self, data: str) -> List[Dict]:
        f = open(data, encoding="gbk", errors='replace')
        cont = f.read()
        f.close()

        df = pd.read_csv(
            StringIO(cont),
            delimiter="\t",
            dtype=self._trader.config.GRID_DTYPE,
            na_filter=False,
        )
        return df.to_dict("records")
