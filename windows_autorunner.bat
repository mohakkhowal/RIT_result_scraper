:start
cls
echo ********** Installing Dependencies **********
set INPUT= "dependencies.txt"
echo INPUT
py -m pip install -r %INPUT%
python bot.py
pause
