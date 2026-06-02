@echo off
if not exist outputs mkdir outputs
echo stop>outputs\stop_collection.flag
echo Stop requested. The app will stop after the current page or category finishes, and completed data will be kept.
pause
