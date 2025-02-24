# ExcelToSQL: A Python Tool for Generating SQL Queries from Excel Files

**Last Updated:** February 20, 2025  
**Author:** (Insert relevant information, if applicable)  

---

## Introduction  

The `ExcelToSQL` tool is a Python-based program that reads data from Excel files (`.xlsx`) and converts them into SQL queries for database operations. It supports **Insert** and **Update** operations, enabling its users to streamline database interactions.

This tool is especially helpful for automating tasks that involve transforming structured Excel data into SQL queries.

---

## Features  

1. **Dynamic Insert and Update Queries:**  
   - Automatically generates SQL commands to insert or update rows of data in a database.  

2. **Flexible User Input:**  
   - The tool allows the user to specify the file name, table name, and columns for operations during runtime.  

3. **Error Handling Options for Invalid Inputs:**  
   - Prompts users and guides them if any required inputs are missing.  

4. **Customizable Output File:**  
   - SQL queries are written to a text file (`sql.txt`) for later execution.

---

## Prerequisites  

Before using the tool, ensure you have the following:  

1. **Python 3.x Installed:**  
   - Download it from [python.org](https://www.python.org/).  

2. **Required Python Libraries:**  
   - `pandas`: For reading and processing Excel files.  
   - Install it by running:  
     ```
     pip install pandas
     ```
   - `openpyxl`: Required for working with `.xlsx` files.  
     Install it by running:  
     ```
     pip install openpyxl
     ```

---

## How to Use  

### Step 1: Clone or Download the Code  
Download the script and save it as `ExcelToSQL.py` in a working directory.

### Step 2: Run the Script  
Execute the script using Python:  
```
python ExcelToSQL.py
```

### Step 3: Complete the Prompts  
The tool will prompt you for the following inputs:  

1. **Operation Type:**  
   - **Enter `u`** for update operations.  
   - **Enter `i`** for insert operations.  

2. **Excel File Name:**  
   - Provide the file name of the input Excel file (e.g., `data.xlsx`).  

3. **Target Table Name:**  
   - Specify the table in the database where the operations will be performed.  

4. **For Update Operations:**  
   - Specify the reference column (key column).  
   - Specify the column(s) to update (comma-separated if multiple).  

### Step 4: Check the Output  
Generated SQL queries will be written to a file named `sql.txt` in the current directory. You can then use this file to run the queries in your database system.

---

## Code Overview  

### Main Methods:  

1. **Initialization (`__init__`):**  
   - Initializes variables and starts the process.  

2. **`TakeUserRequirements`:**  
   - Gathers user inputs, such as file name, operation type, and table name.  

3. **`CheckWhichOperationToRun`:**  
   - Determines which method to execute (`Insert` or `Update`) based on user input.  

4. **Update-Related Functions:**  

   - `UpdateOperation`: Prompts for additional details like reference column and updating columns.  
   - `GenerateSQLForUpdate`: Creates SQL `UPDATE` queries for each line of data extracted from Excel.  
   - `SingleLineUpdateQuery`: Formats the query for each row.  

5. **Insert-Related Functions:**  

   - `MiddleQueryAndEndQueryGenerator`: Constructs the column list (`INSERT INTO tablename (col1, col2)`) and the corresponding values (`VALUES (val1, val2)`).  
   - `GenerateValuesToUpdate`: Converts data into a structure suitable for query generation.  
   - `SingleLineInsertQuery`: Creates a formatted `INSERT` query.  

---

## Error Handling  

1. **Column Not Found (KeyError):**  
   - If a column specified by the user does not exist in the Excel file, the program guides the user to re-enter the correct column names.  

2. **Missing Inputs:**  
   - If a required input is missing (e.g., file name or table name), the program re-prompts the user to fill in the details.  

3. **Invalid Excel File:**  
   - Ensure the Excel file is in `.xlsx` format and readable using the `pandas` library.  

---

## Example SQL Output  

For illustration, assume the tool is pointed to the following Excel data:  

| ID   | Name     | Age |  
|------|----------|-----|  
| 1    | Alice    | 25  |  
| 2    | Bob      | 30  |  

### Update Query Example:  

An update on `employees` table where reference column is `ID`, and the `Name` column is modified:  

```sql
UPDATE EMPLOYEES SET Name='Alice' WHERE ID=1;  
UPDATE EMPLOYEES SET Name='Bob' WHERE ID=2;  
```

### Insert Query Example:  

An insert query for the same table:  

```sql
INSERT INTO EMPLOYEES(ID,Name,Age) VALUES (1,'Alice',25);  
INSERT INTO EMPLOYEES(ID,Name,Age) VALUES (2,'Bob',30);  
```

---

## Troubleshooting Tips  

1. **File Not Found:**  
   - Ensure the Excel file is in the same directory as the program, or provide the full file path.  

2. **Permission Denied:**  
   - Check if the target text file (`sql.txt`) is open in another program and close it.  

3. **Library Import Errors:**  
   - Ensure all required libraries (`pandas`, `openpyxl`) are installed correctly using `pip`.  

4. **Excel Formatting Issues:**  
   - Ensure that the first row of the Excel file contains column headers.  

---

## Enhancements for Future  

- Add support for other SQL operations (DELETE, SELECT).  
- Incorporate GUI-based input for usability.  
- Include more robust error handling and logging mechanisms.  

---

## Contact  

For questions, suggestions, or bug reports, reach out via (Insert preferred communication channel).  

Enjoy using the ExcelToSQL tool! 🎉  
