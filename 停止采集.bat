@echo off
if not exist outputs mkdir outputs
echo stop>outputs\stop_collection.flag
echo 已请求停止采集。程序会在当前页面或当前小类结束后停止，并保留已完成的数据。
pause
