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


https://github.com/user-attachments/assets/edfa4828-e795-43aa-aa13-279e4cf71de1




## Usage

1. Run the bot in UiPath by extracting the repository and open as local project by pressing the json file in the repository.
3. Ensure the Excel file containing contact details is available (you can use the example file).
4. Select the desired method of data extraction (open or closed Excel file).
5. make sure outlook is open, and if you use `Excel file open` make sure the `contact.xlsx` is open
6. The bot will automatically fill in the contact details in Outlook 365.

This bot provides a reliable and efficient way to automate contact creation while maintaining flexibility for different workflows.
# Some insights:
- One of the most difficult things I encountered was identifying buttons and fields and making them generic. The method is to use `AA` for older Windows versions and `UIA` for newer versions. Sometimes the default does not correctly identify the selector.
- It is important to make the code for the selector generic. In the UIA method, the machine gives each element a number or several identifiers. Therefore, in these cases, we use * as a placeholder instead of a series of characters, for example:
```
<uia automationid='app' role='‏‏קבוצה' />
<uia cls='___o3c7u90 f22iagw f122n59 f17wyjut f1mtd64y f1vx9l62 f1c21dwh f*' name='left-rail-appbar' />
<uia cls='___8je7xm0 f10pin' name='אנשים' />
<uia automationid='355fbd79-3ba2-4554-8f2d-0300fde91f30' cls='fui-Button r1alrhcs ___mi33tk0 f1brlhvm f136y8j8 f10xn8zz f1sl3k*' name='אנשים' />
```

