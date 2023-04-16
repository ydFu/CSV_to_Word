# CSV_to_Word
This Python script is designed to read data from a CSV file and populate a Word document with the data. The script uses the csv and docx libraries to read the CSV file and manipulate the Word document, respectively.

## Requirements
 - Python 3.6 or later
 - csv library
 - docx library

## Install

```shell
pip install pandas
pip install python-docx
```

## Usage
  1. Place the CSV file containing the data to be inserted into the Word document in the same directory as the Python script.
  2. Modify the placeholders in the Word document to match the column headers in the CSV file. The placeholders should be in the format &&{column_header}, where {column_header} is the header of the column in the CSV file.
  3. Update the file path of the Word document in the Python script to match the path to the Word document.
  4. Update the file path of the output Word document in the Python script to specify where the output Word document should be saved.
  5. Run the Python script `python CSV_to_Word.py`.

## Example
An example CSV file and Word document have been included in the repository. The CSV file contains data for three different loan types, and the Word document has placeholders for the loan type, interest rate, and loan amount. Running the Python script with the included CSV file will populate the Word document with the data from the CSV file and save the output to a new Word document.


## License
This project is licensed under the MIT License - see the LICENSE file for details.