# Work Item Extract

The project consists in read a CSV file downloaded from an specific query from Azure Devops and treat that data and export this treatment to xlsx file.

The purpose of this project in my case is get the whole data from a project in Azure Devops and try to create a dashboard on Power BI and bring maybe some interesting visualizations and metrics about the project to the leadership.

# About the Query

My query:

![alt text](/images/query.png)

The columns:

![alt text](/images/columns.png)

On top right, click on the three points and export the CSV

![alt text](/images/export.png)

# How it works

Put the downloaded CSV file on the upload_csv directory and run the program. The data will be treated and will generate a XLSX file on xlsx_file directory.

OBS: If there is more than one CSV file in upload_csv, this will not work, you will need to remove the other files and leave only the file to be processed

# Prerequisites

Python 3.11.6 

Pip 23.2

PS: Maybe works with others versions but above is the versions that I used

# How to run

1. Git clone the project

```
git clone git@github.com:matheusfsantana/work-item-extract.git
```
2. Create virtualenv (if you want)

Open the terminal on cloned project directory

```
python3 -m venv venv
```

3. Activate the virtual env

```
source venv/bin/activate #LINUX
``` 

OR

```
venv\Scripts\activate #WINDOWS
```

4. Install requirements

```
pip install -r requirements.txt
```

5. Put the csv file on upload_csv directory and run the application (with venv activated)

```
python main.py
```