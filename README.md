# Google Sheets Data Validation and Form Submission Script

I wanted to collect internal complaints in the company I work at for a RCA project. However, I encountered a limitation with Google Forms, as it couldn't support three layers of dependent dropdown lists. As a result, I took the initiative to design my own form using Google Sheets and AppScript and here how it goes.

---

![image](https://github.com/Mo-Sharabassy/googleform-on-googlesheets/assets/127991953/d7a3338c-3921-4f55-b215-04ec31e080c1)

![image](https://github.com/Mo-Sharabassy/googleform-on-googlesheets/assets/127991953/74cdb53f-de13-4faf-8615-3194f84d33db)


This script is designed to work with Google Sheets and is intended for adding data validation to specific cells in a sheet and for submitting form data. It also includes functions for clearing cells, and sending email notifications.

## Overview

The script serves multiple purposes:

1. **Data Validation**: It sets up ddata validation for specific cells in your Google Sheets document, allowing you to create three-levels dependent dropdown lists based on the selected values in other cells.

2. **Form Submission**: The script enables form data submission. Person1 can input data into the form, and when submitted, it performs various actions, such as saving the data in tab "Sent" and in Person2's Sheet tab "Recieved", sending email notifications to Person2, and clearing the form.

## Usage

Before using the script, you should have a Google Sheets document with two sheets:

- `Form`: This sheet contains the form where users input data.
- `Data`: This sheet contains options with three columns that the form references for data validation. The data for the first and second column are repeated as many times as the third columnd.

The script's primary functions are as follows:

### Data Validation

- **Primary Selection**: When you edit a cell in column E (column 5) of row 14 "Client" in the `Form` sheet, the script applies data validation to another cell based on the selected value. The options for the primary selection are specified in the `Data` sheet.

- **First-level Dependent Selection**: When you edit a cell in column E (column 5) of row 17 "Supplier" in the `Form` sheet, the script applies data validation to another cell based on the selected value. The options for the first-level dependent selection are derived from the `Data` sheet.

- **Second-level Dependent Selection**: When you edit a cell in column E (column 5) of row 20 "Driver" in the `Form` sheet, the script applies data validation to another cell based on both the primary and first-level selections. The options for the second-level dependent selection are derived from the `Data` sheet.

### Form Submission

- Users can input data into the form in the `Form` sheet.

- The script validates the input and checks if all required fields are filled.

- If all fields are filled, it saves the data to the Person1 "Sent" sheet and in Person2 "Received" sheet, sends email notifications to Person2, and clears the form.

### Functions

The script contains the following functions:

- `onEdit(e)`: The main function that responds to cell edits and applies data validation.
- `clearValidationAndContent(range)`: Clears the content and data validation rules for a specified range.
- `applyFirstLevelValidation(val)`: Applies data validation to the first-level dependent selection based on the selected value in the primary selection.
- `applySecondLevelValidation(val)`: Applies data validation to the second-level dependent selection based on the selected values in the primary and first-level selections.
- `applyValidationToCell(list, cell)`: Creates a data validation rule with a list of options and applies it to a specific cell.

Additionally, the script includes functions for clearing specific cells in the form and submitting form data. 

- `clearCells()`: Clears specific cells in the "Form" sheet.
- `submitForm()`: Validates and submits form data, saving it in the "Sent" and "Received" sheets and sending email notifications.
- `sendEmail(email, subject, body)`: Sends email notifications using the MailApp service.

## Installation

To use this script, follow these steps:

1. Create Google Sheets documents.
2. Click on `Extensions` in the top menu.
3. Select `Apps Script`.
4. Replace the default code in the script editor with the code from this repository.
5. Update the URL in submitting_form.js with Person2 Google Sheets after duplicating.
5. Save the project with a name.
6. You can now close the script editor.

The script will automatically respond to cell edits and enable form submission in your `Form` sheet as described above.

## License

This script is provided under the MIT License. See [LICENSE.md](LICENSE.md) for details.

## Contribution

Feel free to contribute to the script by creating issues, making improvements, suggesting alternatives or adding new features.