# -*- coding: utf-8 -*-
import re
import tempfile
import time

import pywinauto
import pywinauto.clipboard

from easytrader import clienttrader, helpers


class HTonClientTrader(clienttrader.BaseLoginClientTrader):
    @property
    def broker_type(self):
        return "hton"

    def login(self, user, password, exe_path, comm_password=None, **kwargs):
        """
        :param user: 用户名
        :param password: 密码
        :param exe_path: 客户端路径, 类似
        :param comm_password:
        :param kwargs:
        :return:
        """
        if comm_password is None:
            raise ValueError("海通必须设置通讯密码")

        try:
            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=1
            )
        # pylint: disable=broad-except
        except Exception:
            self._app = pywinauto.Application().start(exe_path)

            # wait login window ready
            while True:
                try:
                    self._app.top_window().Edit1.wait("ready")
                    break
                except RuntimeError:
                    pass

            self._app.top_window().Edit1.type_keys(user)
            self._app.top_window().Edit2.type_keys(password)

            self._app.top_window().Edit3.type_keys(comm_password)

            self._app.top_window()["连接最快站点"].click()

            # detect login is success or not
            self._app.top_window().wait_not("exists", 30)

            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=10
            )
        self._close_prompt_windows()
        self._main = self._app.window(title="网上股票交易系统5.0")


class GFClientTrader(clienttrader.BaseLoginClientTrader):
    @property
    def broker_type(self):
        return "gf"

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
                path=self._run_exe_path(exe_path), timeout=1
            )
        # pylint: disable=broad-except
        except Exception:
            self._app = pywinauto.Application().start(exe_path)

            # wait login window ready
            while True:
                try:
                    self._app.top_window().Edit1.wait("ready")
                    break
                except RuntimeError:
                    pass

            self._app.top_window().Edit1.type_keys(user)
            self._app.top_window().Edit2.type_keys(password)
            while True:
                try:
                    code = self._handle_verify_code()
                    self._app.top_window().Edit3.type_keys(code)
                    time.sleep(1)
                    self._app.top_window()["确定(Y)"].click()
                    # detect login is success or not
                    try:
                        self._app.top_window().wait_not("exists", 30)
                        break

                    # pylint: disable=broad-except
                    except Exception:
                        self._app.top_window()["确定"].click()

                # pylint: disable=broad-except
                except Exception:
                    raise
                    pass

            self._app = pywinauto.Application().connect(
                path=self._run_exe_path(exe_path), timeout=10
            )
        self._main = self._app.window(title="核新网上交易系统7.70")

    def _handle_verify_code(self):
        control = self._app.top_window().window(control_id=0x5db)
        control.click()
        time.sleep(0.2)
        file_path = tempfile.mktemp() + ".jpg"
        control.capture_as_image().save(file_path)
        time.sleep(0.2)
        vcode = helpers.recognize_verify_code(file_path, "gf")
        return "".join(re.findall("[a-zA-Z0-9]+", vcode))

    def auto_ipo(self):
        self._switch_left_menus(self._config.AUTO_IPO_MENU_PATH)

        stock_list = self._get_grid_data(self._config.COMMON_GRID_CONTROL_ID)

        if len(stock_list) == 0:
            return {"message": "今日无新股"}
        invalid_list_idx = [
            i for i, v in enumerate(stock_list) if v["可申购数量"] <= 0
        ]

        if len(stock_list) == len(invalid_list_idx):
            return {"message": "没有发现可以申购的新股"}

        # self._click(self._config.AUTO_IPO_SELECT_ALL_BUTTON_CONTROL_ID)
        # self.wait(0.1)

        for row in invalid_list_idx:
            self._click_grid_by_row(row)
        self.wait(0.1)

        self._click(self._config.AUTO_IPO_BUTTON_CONTROL_ID)
        self.wait(0.1)

        return self._handle_pop_dialogs()