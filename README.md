# Excel Duplicate Row Merger
A web application for merging and processing Excel files. This tool allows users to upload an Excel file, merge rows based on a unique identifier, and handle discrepancies. It supports adding hyperlinks from the original file to the new merged file and highlights discrepancies in the output. The processed file is then saved for download.
a# Excel File Merger and Processor

A web application for merging and processing Excel files. This tool allows users to upload an Excel file, merge rows based on a unique identifier, and handle discrepancies. It supports adding hyperlinks from the original file to the new merged file and highlights discrepancies in the output. The processed file is then saved for download.

## Limitations

1. **Data Starting Point:** The script processes data starting from the third row of the Excel file. The first row should contain headers, and the second row should contain questions or other relevant information. Ensure that your Excel file is formatted correctly, with the data beginning from row 3.

2. **Excel File Format:** The script currently supports `.xlsx` files. Ensure that your file is in this format, as other formats such as `.xls` or `.csv` are not supported.

3. **Response ID Column:** The script assumes there is a column labeled "Response ID" in the header row. If your file does not have this column or it is named differently, you will need to adjust the script or modify the column name to match.

4. **Cell Discrepancies:** The script identifies discrepancies in cells where values differ and are separated by a pipe (`|`). Ensure that discrepancies are formatted as expected for proper identification.

5. **Hyperlinks:** Hyperlinks are only copied from the original sheet to the new sheet if they are present. If hyperlinks are missing or improperly formatted in the original file, they will not be included in the output.

6. **File Size and Performance:** Very large files may cause performance issues or exceed browser limitations. Test with smaller files first to ensure functionality before using larger datasets.

## User Guide

1. **Upload an Excel File:**
   - Navigate to the web application interface.
   - Use the file input control to select and upload an Excel file.

2. **Process the File:**
   - Click the "Merge" button to start processing the uploaded file.
   - The application will read the Excel file, merge rows based on the "Response ID" column, and handle discrepancies.

3. **Viewing Status:**
   - The status message at the top of the interface will indicate the progress of the processing.
   - You will see a message confirming whether the file was processed successfully or if an error occurred.

4. **Download the Processed File:**
   - Once processing is complete, a download will start automatically, or you will be prompted to save the file.
   - The processed file will be named `merged_output.xlsx` and will include the merged data and a discrepancies sheet.

5. **Handling Errors:**
   - If you encounter an error message, check that the Excel file is formatted correctly and that it includes the necessary columns and data.
   - Ensure that you are uploading a `.xlsx` file and not another format.

6. **Editing the Script (Optional):**
   - If you need to adjust the script for different column names or formats, you can modify the JavaScript code accordingly.
   - For more details on modifying the script, refer to the code comments and documentation within the repository.
