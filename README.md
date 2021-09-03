# Desafio Tunts

## Dependencies:

* pip install gspread
* pip install oauth2client

## Executing:

    You just need to open the .py file with the .json in the same folder.

## Documentation:

### Made using:
* Python 3.9.4 
* Windows 10 64bit
* gspread 4.0.0
* oauth2client 4.1.3

### Functionality:

    Program consists in accessing a google spreadsheet gattering data and using it for calculate the situation of 24 different students,
    after that, it updates the sheet with the situation and grades needed to be approved (if possible and necessary).

    First of all it gets the credentials to use google APIs from the .json file (using a service account created previously in Google Cloud Platform),
    then the program uses the gspread library to get access to the sheet using part of the URL, it proceeds to get the information needed and stores it
    in separete lists, simple operations are then used to calculate everything and store it in sperate lists again, finally, using a function from gspread
    in a simple loop, it updates the columns in the specified range with the information.

    Python was used because of the simplicity and the existing of the previously created library gspread, making the funcitionality and writing much simpler
    and compact.