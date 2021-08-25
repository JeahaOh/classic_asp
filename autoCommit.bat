clear

SET yyyymmdd=%DATE:~0,4%%DATE:~5,2%%DATE:~8,2%

git status

git add .

git commit -m '%yyyymmdd%'

git push

::shutdown -s -t 300

@echo PROCESS END. BYE!