import logging
import os
import time

logger = logging.getLogger()

if not os.path.exists("log"):
    os.mkdir("log")

logger = logging.getLogger()
logger.setLevel(logging.INFO)
ch = logging.StreamHandler()
fh = logging.FileHandler(filename=f"log/log_{time.strftime('%Y%m%d')}.txt")

formatter = logging.Formatter(
    "%(asctime)s - %(module)s - %(funcName)s - line:%(lineno)d - %(levelname)s - %(message)s"
)
formatter2 = logging.Formatter(
    "%(asctime)s -line:%(lineno)d - %(levelname)s - %(message)s")
fh.setFormatter(formatter)
ch.setFormatter(formatter2)
logger.addHandler(ch)  # 将日志输出至屏幕
logger.addHandler(fh)  # 将日志输出至文件