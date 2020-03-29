# e2i Automation Project

This python script converts a multiple tab excel file in to a single tab excel file. 

There has been some changes in the requirements from e2i side. This README file only covers the new version of the whole project. 

ps: [link to the outdated readme](https://github.com/ulaladungdung/e2i-automation/blob/master/README(Outdated).md)

## Pre requisite

Software:

1. Anaconda 
   - Can be installed through this [link](https://www.anaconda.com/distribution/#download-section) **Python 3.7 is recommended**

## List of scripts

1. Webform consolidation
   - Consolidating webforms submitted from users into a single excel file for further processing and auditing.
   - Data from webforms are forwarded to a staff's email, then the script helps automate the collation process.
2. Blacklist comparison
   - Comparing a list of user based on their ID to a list of blacklisted users.
   - The marked users will be consolidated into a single excel form for further processing. 

## Usage

1. Git clone the repository **OR** download the zip file of the repository [here](Downloads\Compressed\e2i-automation-master.zip).

2. Extract the zip file. *(If you use git clone you do not have to extract anything)*

3. Open Anaconda Navigator.

   Type `Anaconda Navigator ` in the windows search.

2. Choose JupyterLab.
   - You will be redirected to your browser (after around 1-2 minutes)
   - The page that is opened on your browser is JupyterLab.
   - **Don't close this page** as you will need this to use your scripts
3. On the left pane, Find the project folder.
4. Choose the automation folder that you want to use.
5. Open the notebook file *(ends with xxxxx.ipynb)* by double clicking it.
   - The 'Cell' that has '**LOCKED**' written is the code that is necessary for the automation to run
   - You only need to change the filenames required in the **non - LOCKED** cell for the automation to run correctly
