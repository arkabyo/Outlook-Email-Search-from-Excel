# Email Subject and Date Extractor: Excel & Outlook - Macro (Using Unique Search Query)

## Overview
This Excel VBA macro module enhances productivity by searching through a designated Outlook email folder for messages that match each selected cell's content in Excel. It then logs the subject and sent date of the first found email back into Excel, facilitating quick reference and data management tasks.

## Features
- **Selective Search**: Performs searches based on the content of user-selected cells in a specific column.
- **Dynamic Column Creation**: Automatically adds "Email Subject" and "Email Date" columns to the Excel sheet if they do not exist.
- **Customizable Email Folder**: Allows searching within the Inbox or a specified subfolder for flexibility in managing diverse email organization structures.
- **User Notifications**: Alerts the user upon completion of the task, enhancing the interactive experience.

## Use Cases
- **Email Management**: Quickly find and log when and regarding what correspondences were made with specific entities or subjects.
- **Audit and Reporting**: Aid in compliance, auditing, and reporting activities by providing an automated way to reference email communications.

## How to Use
1. Open the VBA Editor in Excel and import the module.
2. Ensure Outlook is properly configured and accessible.
3. Select the cells in Column A that contain the search queries.
4. Run the macro and wait for the notification that the task has been completed.

## Time Complexity
The script's performance is subject to its time complexity of `O(n*m)`, where:

- `n` is the number of selected cells in Excel.
- `m` represents the number of emails in the specified Outlook folder.

Given this relationship, the script's execution time will increase with the size of `n` and `m`. To ensure optimal performance, it is recommended to limit the number of selected cells (queries) in each iteration. This approach is particularly beneficial in scenarios involving large volumes of emails, as it helps manage script execution time and resource utilization effectively.


## Collaboration
Feedback and contributions are welcome. Please feel free to fork, submit pull requests, or open issues to discuss potential improvements or features.

