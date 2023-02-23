import time

from wechat import send_wechat

start = time.time()
send_wechat(wechat='傳輸', text='成功!')
print(time.time() - start)
