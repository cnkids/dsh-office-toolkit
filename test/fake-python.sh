#!/bin/sh
# 测试用的假解释器:探测阶段照常回答,执行阶段按 FAKE_MODE 制造异常现场。
if [ "$1" = "-c" ]; then
  if [ "$FAKE_MODE" = "oldpy" ]; then echo "3 6 3.1.5"; else echo "3 12 3.1.5"; fi
  exit 0
fi
case "$FAKE_MODE" in
  garbage)
    echo "Traceback (most recent call last):" >&2
    echo "  File x, line 1" >&2
    echo "这不是 JSON"
    echo "Traceback (most recent call last):" >&2
    exit 1
    ;;
  big)
    head -c 200000 /dev/zero | tr '\0' 'a'
    ;;
  sleep)
    sleep 5
    ;;
  *)
    echo '{"ok":true,"output":"/tmp/none.xlsx","sheets":[]}'
    ;;
esac
