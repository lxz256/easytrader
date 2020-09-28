# coding:utf-8
import re
import time
from typing import Optional

import pywinauto

from easytrader import exceptions
from easytrader import logger
from easytrader.utils.perf import perf_clock
from easytrader.utils.win_gui import SetForegroundWindow, ShowWindow


class PopDialogHandler:
    def __init__(self, app):
        self._app = app

    def _set_foreground(self, grid=None):
        if grid is None:
            grid = self._trader.main
        if grid.has_style(pywinauto.win32defines.WS_MINIMIZE):  # if minimized
            ShowWindow(grid.wrapper_object(), 9)  # restore window state
        else:
            SetForegroundWindow(grid.wrapper_object())  # bring to front

    @perf_clock
    def handle(self, title):
        if any(s in title for s in {"提示信息", "委托确认", "网上交易用户协议"}):
            self._submit_by_shortcut()
            time.sleep(0.5)
            return None

        if "提示" in title:
            content = self._extract_content()
            self._submit_by_click()
            time.sleep(0.5)
            return {"message": content}

        content = self._extract_content()
        self._close()
        return {"message": "unknown message: {}".format(content)}

    def _extract_content(self):
        return self._app.top_window().Static.window_text()

    def _extract_entrust_id(self, content):
        return re.search(r"\d+", content).group()

    def _select_checkbox(self):
        self._app.top_window().child_window(control_id=1504, class_name='Button').click()

    def _submit_by_click(self):
        # self._set_foreground(self._app.top_window())

        if self._app.top_window()['是(&Y)'].exists():
            self._app.top_window()['是(&Y)'].click()
        elif self._app.top_window()["确定"].exists():
            self._app.top_window()["确定"].click()
        else:
            self._app.top_window()["确认"].click()

    def _submit_by_shortcut(self):
        self._set_foreground(self._app.top_window())
        self._app.top_window().type_keys("%Y", set_foreground=False)

    def _close(self):
        self._app.top_window().close()


class TradePopDialogHandler(PopDialogHandler):
    @perf_clock
    def handle(self, title) -> Optional[dict]:
        content = self._extract_content()
        logger.debug('handling title:%s content:%s' % (title, content))

        if title == "委托确认":
            self._submit_by_shortcut()
            return None

        if title == "提示信息":
            if "超出涨跌停" in content:
                self._submit_by_shortcut()
                time.sleep(0.5)
                return None

            if "委托价格的小数价格应为" in content:
                self._submit_by_shortcut()
                time.sleep(0.5)
                return None

            if "逆回购" in content:
                self._submit_by_shortcut()
                return None

            if "基金申购委托" in content:
                self._submit_by_shortcut()
                return None

            # 银河申购第1个窗口提示信息的Static取不到值（win10和server分别取到如下），暂时处理如下
            if "提示信息" in content or content == '':
                self._submit_by_click()
                return None

            return None

        if title == "提示":
            self._submit_by_click()
            if "成功" in content:
                entrust_no = self._extract_entrust_id(content)
                return {"entrust_no": entrust_no}
            else:
                raise exceptions.TradeError(content)

        # 银河基金信息披露和风险确认
        if title == "基金信息披露":
            self._app.top_window()['基金信息披露Shell DocObject View'].click()
            self._app.top_window().type_keys('{TAB}')
            self._app.top_window().type_keys('{TAB}')
            self._app.top_window().type_keys("{ENTER}")
            # self._app.top_window().child_window(control_id=1504, class_name='Button').wait('ready', timeout=10)

            # retry = 10
            # while retry:
            #     try:
            #         self._app.top_window().child_window(control_id=4427, class_name='Button').wait("ready", timeout=0.5,
            #                                                                                        retry_interval=0.2)  # 保存
            #         break
            #     except RuntimeError:
            #         retry -= 1
            #         logger.info('con not find save button, retry%s ' % (10 - retry))
            #
            # time.sleep(0.2)
            # self._app.top_window().type_keys("{ESC}")
            # self._app.top_window()['保存(&S)'].wait_not('ready', 5)

            self._select_checkbox()
            time.sleep(0.1)
            self._submit_by_click()
            return None

        # 银河风险告知
        if title == "公募证券投资基金投资风险告知":
            self._select_checkbox()
            time.sleep(0.1)
            self._submit_by_click()
            return None
        # 银河适当性匹配
        if title == "适当性匹配结果确认书":
            self._submit_by_click()
            return None

        raise exceptions.TradeError('未知window, title:%s content:%s' % (title, content))
