# gsheet-to-docx
You will need a ```credential.json``` in ```conf``` which is not in the repo and should never be. Get your local copy and never commit it to repo

copy ```conf/config.yml.dist``` as ```conf/config.yml``` and do not commit the copied file

```
cd /home/asif/projects/asif@github/gsheet-to-docx/src
python docx-from-gsheet.py --config "../conf/config.yml"
```
