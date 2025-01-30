# Contact Creation Bot

This project implements a UiPath bot that automates the creation of new contacts in Outlook 365 based on data from an Excel file. The bot supports two different methods:

1. **Using an open Excel file** – The bot extracts data from an already opened Excel file.
2. **Using a closed Excel file** – The bot reads the data from an Excel file without requiring it to be open. using macro VBA script.

## Features

- Designed to work with Outlook 365.
- Ensures a generic and adaptable element selection process for users running this version of Outlook.
- Robust handling of edge cases to improve stability and reliability.
- Efficiently processes contact creation while minimizing potential errors.

## Demonstration

Two demonstration videos are provided to showcase the bot in action, highlighting both methods of execution.

## Usage

1. Run the bot in UiPath by extracting the repository and open as local project by pressing the json file in the repository.
3. Ensure the Excel file containing contact details is available (you can use the example file).
4. Select the desired method of data extraction (open or closed Excel file).
5. The bot will automatically fill in the contact details in Outlook 365.

This bot provides a reliable and efficient way to automate contact creation while maintaining flexibility for different workflows.

