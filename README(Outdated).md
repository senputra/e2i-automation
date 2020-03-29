# e2i Automation Project

This python script converts a multiple tab excel file in to a single tab excel file. 

## Pre requisite

Software:

1. Python
2. Pandas module in python



## Usage

**Steps:**

1. Download the zip file from this link <https://codeload.github.com/ulaladungdung/e2i-automation/zip/master>
2. Extract the zip file onto a folder
3. Open Command Prompt
   1. On the Explorer inside the folder with the extracted files
   2. Hold `Shift` then `right click`
   3. Choose `Open PowerShell window here`
4. Copy the Excel Book that needs to be converted into the folder
5. Type in `python ./automate.py -i ./input_file_name.xlsx -o ./output_file_name.xlsx`



### Use case #1 (Input file is given and output file is set to default)

on `line 5` above use

```bash
$ python ./automate.py -i ./input_file_name.xlsx
```

The output file will be set to the default file `Consolidated_data.xlsx` in the same folder.

The output file will be `output_file_name.xlsx`. 

If the file **exists**, it will add new rows under the existing excel

If the file **does not exists**, it will create a new file and append the new rows in to the excel file

**Download**

```bash
# Clone github repository
$ git clone https://github.com/ulaladungdung/e2i-automation.git
$ python ./automate.py -i ./inputdata.xlsx -o ./outputdata.xlsx
```

**Arguments**: 

- `-o` output file path
- `-i` input file path

**Note**:

​      The output file will be overwritten.

