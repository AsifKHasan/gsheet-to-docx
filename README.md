## gsheet-to-docx
# Software and Tools required
1. Python 3.8 or higher - https://www.python.org/downloads/
2. Git -  https://git-scm.com/downloads


# Verify the installations
1. create a directory in your laptop where you keep your office files and programs. Normally I keep them all under D:\projects. Chose your own directory
2. open a command prompt and cd to D:\projects (or whatever your chosen directory is)
3. run command ```python -version``` and send me the output
4. run command ```python -m pip install --upgrade pip``` and send me the output
5. run command ```pip --version``` and send me the output
6. run command ```git --version``` and send me the output


# Get the scripts/programs
1. cd to d:\projects
2. run ```git clone https://github.com/AsifKHasan/gsheet-to-docx.git``` and send me the output
3. cd to D:\projects\gsheet-to-docx
4. run command ```pip install -r requirements.txt```. See if there is any error or not. If you get errors share the output with me.


# Configure the scripts/programs
You will need a ```credential.json``` in ```conf``` which is not in the repo and should never be. Get your local copy and never commit it to repo

1. get a file named *credential.json* from me and paste/copy to D:\projects\gsheet-to-docx\conf
2. copy the file D:\projects\gsheet-to-docx\conf\config.yml.dist as a new file D:\projects\gsheet-to-docx\conf\config.yml

or

3. copy ```conf/config.yml.dist``` as ```conf/config.yml``` edit it and do not commit the copied file


# Running scripts/programs
```
cd /home/asif/projects/asif@github/gsheet-to-docx/src
python docx-from-gsheet.py --config "../conf/config.yml"
```
