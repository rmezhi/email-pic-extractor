# Email Picture Extractor

This Google Apps Script is designed to automate the process of downloading images and other files from specific Gmail emails to your Google Drive. It searches for emails based on a subject filter, extracts download links, and handles various file types, including images and zip archives.

## Key Features

- **Gmail Search**: Searches for emails using a specific subject line and filters out already processed emails using a "Processed" label.
- **Link Extraction**: Extracts all download links from the email body.
- **Redirect Handling**: Follows redirect URLs to get the final download link.
- **File Type Verification**: Checks the `Content-Type` of the linked file to determine if it's an image or a zip archive.
- **Image Downloading**: Downloads and saves image files to a specified Google Drive folder.
- **Zip File Handling**: Downloads zip archives, unzips them, and saves the contents to the same Google Drive folder.
- **Duplicate Prevention**: Checks if a file with the same name already exists in the target folder to avoid duplicates.
- **Error Handling**: Includes robust `try-catch` blocks to handle potential errors during execution, such as network issues, invalid links, or permission problems.
- **Logging and Summary**: Provides detailed logging throughout the script's execution and a summary at the end, including the number of processed messages and downloaded files.
- **Automated Labeling**: Marks processed emails with a "Processed" label to prevent them from being processed again.

## Setup Instructions

1.  **Open Google Apps Script**: Go to [script.google.com](https://script.google.com).
2.  **Create a New Project**: Click on "New project" and give it a name (e.g., "Email Pic Extractor").
3.  **Copy the Script**: Copy the code from `download_images.gs` and paste it into the script editor.
4.  **Save the Project**: Click the save icon or press `Ctrl+S` (`Cmd+S` on Mac).
5.  **Run the Script**:
    *   Select the `downloadImagesFromEmails` function from the dropdown menu next to the "Debug" button.
    *   Click "Run".
6.  **Grant Permissions**: The first time you run the script, you will be prompted to grant permissions for it to access your Gmail and Google Drive. Follow the on-screen instructions to authorize the script.

## Configuration

You can customize the script by modifying the following constants at the top of the `downloadImagesFromEmails` function:

- `subjectFilter`: The subject line of the emails you want to process.
- `folderName`: The name of the folder in your Google Drive where the files will be saved.
- `processedLabelName`: The name of the Gmail label to apply to processed emails.

## Usage

Once the script is set up and configured, you can run it manually from the Apps Script editor. For automated execution, you can set up a time-driven trigger:

1.  In the Apps Script editor, click on the "Triggers" icon (it looks like a clock).
2.  Click "Add Trigger".
3.  Configure the trigger to run the `downloadImagesFromEmails` function at your desired frequency (e.g., daily, weekly).

## Error Handling and Logging

The script includes comprehensive error handling and logging. To view the logs:

1.  Run the script from the Apps Script editor.
2.  After execution, go to "View" > "Logs" (`Ctrl+Enter` or `Cmd+Enter`) to see the detailed output.

This will help you troubleshoot any issues that may arise.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

