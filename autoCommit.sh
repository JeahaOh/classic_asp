  
echo -e '\r-- GIT STATUS --\n\r'
git status

echo -e '\r-- GIT PULL --\n\r'
git pull

echo -e '\r-- GIT ADD . --\n\r'
git add .

echo -e "\r-- GIT COMMIT $(date +%Y%m%d) --\n\r"
git commit -m "$(date +%Y%m%d)"

echo -e "\r-- GIT PUSH --\n\r"
git push