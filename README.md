# Google Sheets Mail Merge

This project is a Google Sheets Mail Merge script that allows you to send customized emails using draft emails from your Gmail account. You can send the emails immediately or schedule them for a future date and time.

## Prerequisites

To use this script, you will need:

1. A Google account with access to Google Sheets and Gmail
2. A Google Sheets document with the recipient data
3. A Gmail account with at least one draft email containing placeholders for the data you want to merge

## Installation

1. Open the Google Sheets document where you want to use the mail merge script.
2. Click on `Extensions` in the menu, then select `Apps Script`.
3. Delete any existing code in the `Code.gs` file and replace it with the contents of the `code.gs` file from this repository.
4. Create a new HTML file in the Apps Script project, name it `MailMerge.html`, and replace its contents with the contents of the `MailMerge.html` file from this repository.
5. Replace the `appscript.json` contents in your Apps Script project with the contents of the `appscript.json` file from this repository.

## Obtaining your API Key and Client ID

To use the script, you will need a Google API Key and a Google Client ID. Follow these steps to obtain them:

1. Go to the [Google API Console](https://console.developers.google.com/).
2. Create a new project or select an existing one.
3. In the Dashboard, click on `ENABLE APIS AND SERVICES`.
4. Search for "Gmail API" and "Google Sheets API" and enable them for your project.
5. Click on `Credentials` in the left sidebar.
6. Click on `+ CREATE CREDENTIALS` and select `API key`. Copy the generated API key and replace the placeholder `YOUR_API_KEY` in the `code.gs` file.
7. Click on `+ CREATE CREDENTIALS` again and select `OAuth client ID`. Choose `Web application` as the application type, and enter a name for the client. Copy the generated client ID and replace the placeholder `YOUR_CLIENT_ID` in the `code.gs` file.

## Creating a Template from Draft Email

To create a template for your mail merge, you'll need to use Gmail's drafts feature. Follow these steps to create a template using draft emails:

1. Compose a new email in Gmail.
2. Use placeholders for the dynamic content you want to include in your email. The placeholders should be wrapped in double curly braces, like `{{header}}`, `{{header2}}`, etc.
3. Save the email as a draft by closing the email composer without sending it.

For example, if your Google Sheet has columns with headers "First Name" and "Last Name", you can create a draft email with the following content:

> Subject: Welcome, {{First Name}}!
>
> Hi {{First Name}} {{Last Name}},
>
> Welcome to our community! We are thrilled to have you with us.
>
> Best regards,
>
> Your Team


When the mail merge runs, it will replace the `{{First Name}}` and `{{Last Name}}` placeholders with the actual data from the corresponding cells in the Google Sheet.

Remember to save the email as a draft and not to send it. The mail merge script will use the draft as a template when sending out the emails.

## How to Use the Mail Merge Script

1. Save and deploy the Apps Script project by clicking `File > Save` and `Publish > Deploy as web app` in the Apps Script editor.
2. Return to your Google Sheets document and refresh the page.
3. A new menu item called `Mail Merge` should now appear in the menu bar. Click on it and select `Start Mail Merge`.
4. A dialog will open with the following options:
   - Recipient Column: Select the column in the sheet containing the recipient email addresses.
   - Draft Email: Select the Gmail draft email you want to use as the template for the mail merge.
   - Option: Choose whether to send the emails immediately or schedule them for a later date and time.
   - Date and Time: If scheduling, choose the date and time for the emails to be sent.
5. Click on `Start Mail Merge` to begin the mail merge process. The script will send or schedule the emails based on your selections, and update the Google Sheet with the merge status.

**Note:** The scheduled emails are limited to a maximum of 20 per project due to Google Apps Script trigger limitations. You can delete old schedules by clicking the `Delete Old Schedule` button in the Mail Merge dialog.

## Features

:heavy_check_mark: Send multiple row

:heavy_check_mark: Schedule by date and time

:heavy_check_mark: Placeholder on subject line

## Future Roadmap

:white_large_square: 1. Add batch scheduling like YAMM to schedule more than 20 emails  
:white_large_square: 2. Add progress bar  
:white_large_square: 3. Add custom email tracking (opened or not)  
:white_large_square: 4. Templates Gallery: Create a gallery of customizable email templates for various use cases  
:white_large_square: 5. Attachments Support: Allow users to include file attachments from Google Drive or local storage  
:white_large_square: 6. A/B Testing: Implement A/B testing functionality for optimizing email content  
:white_large_square: 7. Analytics Dashboard: Develop a dashboard to display key email performance metrics  
:white_large_square: 8. Conditional Formatting: Introduce conditional formatting in email templates  
:white_large_square: 9. Auto Follow-up: Enable automatic follow-up emails for non-responsive recipients  
:white_large_square: 10. Integration with Other Google Workspace Apps: Integrate with Google Calendar, Google Forms, and other apps
:white_large_square: 11. Add append cc to sheet