# -*- coding: utf-8 -*-
import re
import tempfile

import pywinauto

from easytrader import clienttrader, grid_strategies, pop_dialog_handler
from easytrader.utils.captcha import recognize_verify_code
from easytrader.utils.perf import perf_clock


class YHClientTrader(clienttrader.BaseLoginClientTrader):
    """
    Changelog:

    2018.07.01:
        银河客户端 2018.5.11 更新后不再支持通过剪切板复制获取 Grid 内容，
        改为使用保存为 Xls 再读取的方式获取
    """

    grid_strategy = grid_strategies.Xls

    @property
    def broker_type(self):
        return "yh"

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
        登陆客户端
        :param user: 账号
        :param password: 明文密码
        :param exe_path: 客户端路径类似 'C:\\中国银河证券双子星3.2\\Binarystar.exe',
            默认 'C:\\中国银河证券双子星3.2\\Binarystar.exe'
        :param comm_password: 通讯密码, 华泰需要，可不设
        :return:
        """
        try:
            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path)
            )
        # pylint: disable=broad-except
        except Exception:
            self._app = pywinauto.Application().start(exe_path)
            is_xiadan = True if "xiadan.exe" in exe_path else False

            # wait login window ready
            self._app.top_window().Edit1.wait("ready", 100)
            self._app.top_window().Edit1.type_keys(user)
            self._app.top_window().Edit2.type_keys(password)
            while True:
                self._app.top_window().Edit3.type_keys(
                    self._handle_verify_code(is_xiadan)
                )
                self._app.top_window()["委托登录"].click()

                # detect login is success or not
                try:
                    self._app.top_window().wait_not("exists visible", 10)
                    # break
                except Exception:
                    self._app.top_window()["确定"].click()  # 验证码
                    continue

                self._app = pywinauto.Application().connect(
                    path=self._run_exe_path(exe_path), timeout=10
                )
                self._main = self._app.window(title="网上股票交易系统5.0")
                try:
                    self._main.child_window(
                        control_id=129, class_name="SysTreeView32"
                    ).wait("ready", 30)
                    break
                # pylint: disable=broad-except
                except Exception:
                    self._app.top_window()["确定"].click()  # 柜台无响应
                    self.wait(60*5)
                    continue
        self._close_prompt_windows()

    def _switch_window_to_normal_mode(self):
        self._app.top_window().child_window(
            control_id=32812, class_name="Button"
        ).click()

    def _handle_verify_code(self, is_xiadan):
        control = self._app.top_window().child_window(
            control_id=1499 if is_xiadan else 22202
        )
        control.click()

        file_path = tempfile.mktemp()
        if is_xiadan:
            rect = control.element_info.rectangle
            rect.right = round(
                rect.right + (rect.right - rect.left) * 0.3
            )  # 扩展验证码控件截图范围为4个字符
            control.capture_as_image(rect).save(file_path, "jpeg")
        else:
            control.capture_as_image().save(file_path, "jpeg")
        verify_code = recognize_verify_code(file_path, "yh_client")
        return "".join(re.findall(r"\d+", verify_code))

    @property
    def balance(self):
        self._switch_left_menus(self._config.BALANCE_MENU_PATH)
        return self._get_grid_data(self._config.BALANCE_GRID_CONTROL_ID)[0]

    def auto_ipo(self, kcb_flag=0):
        self._switch_left_menus(self._config.AUTO_IPO_MENU_PATH)
        stock_list = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)
        if len(stock_list) == 0:
            return {"message": "今日无新股"}
        invalid_list_idx = [
            i for i, v in enumerate(stock_list) if v["申购数量"] <= 0 or (v["新股代码"]//1000 == 787 and not kcb_flag)
        ]
        if len(stock_list) == len(invalid_list_idx):
            return {"message": "没有发现可以申购的新股"}
        self.wait(0.1)
        # for row in invalid_list_idx:
        # self._click_grid_by_row(row)
        self._click(self._config.AUTO_IPO_BUTTON_CONTROL_ID)
        self.wait(0.1)
        return self._handle_pop_dialogs()

    @perf_clock
    def purchase(self, security, amount):
        """
        市价交易
        :param security: 六位证券代码
        :param amount: 申购金额
        :return: {'entrust_no': '委托单号'}
        """
        self._switch_left_menus(["场内基金", "基金申购"], sleep=0.2)

        self._type_edit_control_keys(self._config.TRADE_SECURITY_CONTROL_ID, security)
        self._type_edit_control_keys(self._config.TRADE_AMOUNT_CONTROL_ID, str(int(amount)))
        self._submit_trade()

        return self._handle_pop_dialogs(
            handler_class=pop_dialog_handler.TradePopDialogHandler
        )