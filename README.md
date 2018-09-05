# UniandesTeamworkReporter
Reporter for TA's Uniandes Software Development Class

### Dependencies

````
pip install python-docx==0.8.6
pip install git+https://github.com/ravelinx22/python-pptx@feature/update_charts
pip install xlrd==1.1.0
````
Go where your python-docx is install under python2.7/site-packages/ and replace the docx/ folder with the following:

Download: https://www.dropbox.com/s/27pwxtd3mdj9im1/python-docx.zip?dl=0

### Running scripts

#### Generate data

Download excel reports from TeamWork and locate them inside the Excels/ folder of the project. Then cd to the Excels/ folder and run the following command:

````
python generate_data.py FILE_NAME 
````

***Note: You don't need to write the extension .xls in the FILE_NAME***

This will give you the data that is going to be in the report.

#### Generate report

In another terminal cd to the root folder of the project and run the following command:

````
python generate_document.py
````

Insert the data that was generated before. When you finish a .docx document will be generated in the root folder of the project.
