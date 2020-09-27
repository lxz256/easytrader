__author__ = 'liuxz'
__date__ = '2019/12/29'

import time

import easytrader

# user = easytrader.use('htzq_client')
# user.prepare(user='1280183227', password='126259', comm_password='126259')
# time.sleep(0.1)
# user.purchase(security='162411', amount=100)

user = easytrader.use('yh_client')
user.prepare(user='213600010660', password='126259')
time.sleep(1)
# print(user._app.top_window().child_window(control_id=1504, class_name='Button').window_text())
#
user.purchase(security='162411', amount=100)

# # todo 工作日检查
# # todo 溢价检查
# # todo 逐账户申购
# selects = user.main.child_window(
#     control_id=user.config.ACCOUNT_CONTROL_ID, class_name="ComboBox"
# )
# for i in range(len(set(selects.texts()))):
#     selects.select(i)
# o = user.position
# print(o)

# print(user.app.window(class_name="#32770", visible_only=True))
# print(user.app.top_window().child_window(control_id=1365).window_text())
