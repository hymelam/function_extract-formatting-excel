
# How to extract indentation information from an Excel (xlsx) file

Excel files sometimes store meaningful information in cell formatting.

In the example below, rows beginning with "Male" and "Female" contain no information about field of study. The nesting of "Male" and "Female" within field of study (and field of study within "All Fields") is made apparent by indentation.
<br/><br/>
![](excel_indent_example.PNG)
<br/><br/>
(This Excel file was downloaded from: https://ncses.nsf.gov/pubs/nsf19301/data [Table 14])

While this example is simple and could be easily handled during data cleaning, one can easily imagine more complex scenarios in which the nesting structure isn't as obvious (e.g., there are more than two unique strings - male and female - being nested, identical strings existing at multiple levels of the nesting structure, etc.).

The goal of this exercise was to create a function that:
1. Extracts information about indentation from an Excel file (Which cells are indented? How much are they indented?)
2. Returns that information in a matrix whose structure mimics the original Excel file (i.e., if the Excel file contained 100 rows and 5 columns of formatted data, the function should return a 100x5 matrix of numbers where 0 = No indentation, 1 = first level of indentation, 2 = second level of indentation, and so on). This data could be used in later data cleaning.

Additionally, this is my first time using Python to problem-solve, so it serves as a useful learning exercise. I make no claims that this code adheres to any sort of "best practices" - however, one of the easiest ways for me to learn is to dive right in and figure out what I <b>don't</b> know. A list of questions that this exercise raised for me will be included at the end of the notebook.


```python
import numpy as np 
from openpyxl import load_workbook
```


```python
def excel_indent_finder(file, sheet_number= 1):
    
    # Import the Excel Workbook
    wb = load_workbook(file, read_only=True)
    
    # Subtract 1 from worksheet number argument
    # (The first Excel worksheet has an index of 0)
    sheet_index = sheet_number - 1
    
    # Show a custom error message and exit if the worksheet does not exist.
    try:
        wb.worksheets[sheet_index] 
    except IndexError:
        print("Error: Worksheet does not exist. Enter the worksheet number as an interger starting from 1.")
        return
    # If the worksheet does exist...
    else:
        ws = wb.worksheets[sheet_index] 
        # Get the max row and column
        max_row = ws.max_row
        max_col = ws.max_column
        # Return information about the worksheet to the user (name, number of rows and columns)
        print("Returning attributes of worksheet: '" + ws.title + "'")
        print(ws.title + " contains " + str(max_row) + " rows and " + str(max_col) + " columns")

    # Prepare to save information about the indentation formatting
    format_matrix = [];
    # Iterate over the rows
    for row in ws.rows:
        # And while you're in a row, iterate across the cells/columns
        for cell in row:
            # If cell has a value, then look for (and append) the indentation information
            if cell.value: 
                format_matrix.append(cell.alignment.indent)  
            # If cell does not have a value, append a NaN
            else:
                format_matrix.append(np.nan)
    # Reshape the matrix so that it mirrors the original Excel file
    format_reshaped = np.reshape(format_matrix, (max_row, max_col)) 
    # Close the connection to the Excel file (Unsure of this part)
    wb._archive.close()
    # Return the reshaped matrix
    return(format_reshaped)
```


```python
my_format_matrix = excel_indent_finder("sed17-sr-tab014.xlsx")
```

    Returning attributes of worksheet: 'Table 14'
    Table 14 contains 32 rows and 15 columns
    


```python
print(my_format_matrix)
```

    [[ 0. nan nan nan nan nan nan nan nan nan nan nan nan nan nan]
     [ 0. nan nan nan nan nan nan nan nan nan nan nan nan nan nan]
     [ 0. nan nan nan nan nan nan nan nan nan nan nan nan nan nan]
     [ 0.  0. nan  0. nan  0. nan  0. nan  0. nan  0. nan  0. nan]
     [nan  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 1.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]
     [ 2.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.  0.]]
    

The output looks good! The NaNs mark merged and empty cells.


```python
# Save output as a .csv (need to use float instead of integer due to NaNs)
np.savetxt("matrix_output.csv", my_format_matrix, fmt = "%f",delimiter=",")
```

#### Questions and topics for research:
1. Research best practices for error handling. I've included one Try/Catch here as an experiment.
2. Research best practices for import statements. What happens if they are included inside functions? What if the module is already imported? What if "import as" was used? Need to learn where modules live in the environment, in general.
3. Research best practices for building matrices in Python. The method used here, for example, is memory intensive in R. (I wouldn't be surprised if an improvement is to initialize a matrix with the known final dimensions, rather than one that is completely empty.)
4. Experiment with openpyxl to learn about how it handles file locking (of the imported xlsx file).
5. Research how to read in functions from external Python files (like R's `source()`).

#### Ideas for how the function might be extended:
1. Add an argument allowing the user to search for one of multiple types of Excel formatting (e.g., highlighted cells), instead of only indentation.
2. Allow the user to specify the worksheet by either number (currently implemented) or worksheet name
3. Allow customization of what is returned if the cell is empty/merged (currently NaN). 
4. In general, test the function on Excel workbooks containing different components (graphs, images, etc.). At this point, I only know that the function works on this workbook (and similar simple workbooks I've tested outside this Notebook).
