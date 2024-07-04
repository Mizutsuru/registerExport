
# registerExport

registerExport is a simple program made with Python that allows users to create an excel data report using a template and a data file.

The program selects a sheet from the base file (the template) specified in the code, a sheet from the excel file with all the data, and generates a new resulting file from both files.


 


## Variables

To run this project, you will need to edit if necessary the following variables on .py file:

`fullPlantilla` : template document['<sheet_name>']

`fullnou` : <data_file_name>.create_sheet(title="<name_of_new_sheet>")

`delete_sheet` : [fulla for fulla in <data_file_name>.sheetnames if fulla.startswith('<sheet_name>')]. 
- NOTE: delete_sheet is a variable that removes duplicate sheets, in case the program generates the template document sheet multiple times simultaneously.
