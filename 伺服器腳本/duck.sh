#!/bin/bash
# DuckDNS 更新脚本

# ====== 以下两行改成你的值 ======
DOMAIN=onlycolor80193922
TOKEN=51ff4b38-25ef-4407-ad3c-a212f6c267bb
# =================================

# 调用 DuckDNS 的更新接口（自动填当前 IP）
curl -s "https://www.duckdns.org/update?domains=$DOMAIN&token=$TOKEN&ip=" \
  >/dev/null 2>&1
# 加上 verbose 参数，并去掉重定向
curl "https://www.duckdns.org/update?domains=$DOMAIN&token=$TOKEN&ip=&verbose=true"