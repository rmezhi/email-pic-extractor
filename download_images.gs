/**
 * This script searches for Gmail messages with a specific subject, extracts image download links,
 * resolves any redirects, and saves the images to a specified folder in Google Drive.
 *
 * Key Features:
 * - Handles soft line breaks in email bodies.
 * - Resolves redirect URLs to fetch the final image.
 * - Saves images with unique filenames based on timestamps.
 * - Marks messages as read after successful processing.
 * - Labels processed emails to avoid reprocessing.
 * - Skips downloading images if they already exist in the target folder.
 * - Verifies that the target of the link is an image before downloading.
 * - Unzips and saves files from zip archives.
 */

function downloadImagesFromEmails() {
  const subjectFilter = `"Polar Bears (Twos)"`; // Subject filter for Gmail search
  const folderName = "daycare-pics"; // Target folder name in Google Drive
  const processedLabelName = "Processed"; // Label to mark processed emails

  let processedMessages = 0;
  let downloadedFiles = 0;
  let unzippedFilesCount = 0;

  console.log(`Starting script to download images from emails with subject: ${subjectFilter}`);

  try {
    // Get the folder in Google Drive, create it if it doesn't exist
    let targetFolder = getOrCreateFolder(folderName);

    // Get or create the "Processed" label
    const processedLabel = getOrCreateLabel(processedLabelName);

    // Search for emails with the specified subject that do not have the "Processed" label
    const threads = GmailApp.search(`${subjectFilter} -label:${processedLabelName}`);
    console.log(`Found ${threads.length} email threads with subject: ${subjectFilter} and without label: ${processedLabelName}`);

    const messages = GmailApp.getMessagesForThreads(threads);
    console.log(`Processing ${messages.flat().length} messages in total.`);

    messages.flat().forEach((message, index) => {
      try {
        console.log(`Processing message ${index + 1}...`);
        const body = message.getBody();

        if (!body) {
          console.warn(`Message ${index + 1} has an empty body. Skipping.`);
          return;
        }

        const sanitizedBody = body.replace(/=\r?\n/g, ''); // Remove soft line breaks from the email body

        // Extract all download links from the email body using a regex pattern
        const regex = /https:\/\/email3\.kaymbu\.com\/ls\/click\?upn=[^\s"']+/g;
        const links = sanitizedBody.match(regex);

        if (!links || links.length === 0) {
          console.warn(`No valid links found in message ${index + 1}. Skipping.`);
          return;
        }

        console.log(`Found ${links.length} download link(s) in message ${index + 1}.`);
        links.forEach(link => {
          try {
            console.log(`Attempting to process link: ${link}`);
            // Fetch the redirect URL (do not follow redirects automatically)
            const response = UrlFetchApp.fetch(link, { muteHttpExceptions: true, followRedirects: false });

            if (response.getResponseCode() !== 302) {
              console.error(`Unexpected response code ${response.getResponseCode()} for link: ${link}`);
              return;
            }

            const finalImageUrl = response.getHeaders().Location;

            if (!finalImageUrl) {
              console.error(`No redirect URL found in headers for link: ${link}. Full response headers: ${JSON.stringify(response.getHeaders())}`);
              return;
            }

            console.log(`Redirect resolved to final URL: ${finalImageUrl}`);

            // Check if the target is an image by inspecting the Content-Type header
            const validationResponse = UrlFetchApp.fetch(finalImageUrl, { muteHttpExceptions: true });
            const contentType = validationResponse.getHeaders()["Content-Type"] || "";

            if (contentType.startsWith("image/")) {
              // Handle image files
              const urlParts = finalImageUrl.split("/");
              const originalFilename = urlParts[urlParts.length - 1];

              if (fileExistsInFolder(targetFolder, originalFilename)) {
                console.log(`File '${originalFilename}' already exists. Skipping.`);
                return;
              }

              const blob = validationResponse.getBlob();
              targetFolder.createFile(blob.setName(originalFilename));
              console.log(`Successfully downloaded and saved image: ${originalFilename}`);
              downloadedFiles++;
              message.markRead();

            } else if (contentType === "application/zip" || contentType === "application/x-zip-compressed") {
              // Handle zip files
              console.log(`Found a zip file at: ${finalImageUrl}`);
              const zipBlob = validationResponse.getBlob();
              try {
                const unzippedFiles = Utilities.unzip(zipBlob);
                unzippedFiles.forEach(fileBlob => {
                  const filename = fileBlob.getName();
                  if (fileExistsInFolder(targetFolder, filename)) {
                    console.log(`File '${filename}' from zip already exists. Skipping.`);
                  } else {
                    targetFolder.createFile(fileBlob);
                    console.log(`Unzipped and saved file: ${filename}`);
                    unzippedFilesCount++;
                  }
                });
              } catch (unzipError) {
                console.error(`Failed to unzip file from ${finalImageUrl}. Error: ${unzipError.message}`);
              }
              message.markRead();

            } else {
              console.log(`Skipping non-image/zip link: ${finalImageUrl} (Content-Type: ${contentType})`);
            }
          } catch (linkError) {
            console.error(`Error processing link: ${link}. Error: ${linkError.message}`);
          }
        });

        // Add the "Processed" label to the thread
        const thread = message.getThread();
        try {
          thread.addLabel(processedLabel);
          console.log(`Added label '${processedLabelName}' to thread containing message ${index + 1}.`);
          processedMessages++;
        } catch (labelError) {
          console.error(`Failed to add label to thread for message ${index + 1}. Error: ${labelError.message}`);
        }
      } catch (messageError) {
        console.error(`Error processing message ${index + 1}. Error: ${messageError.message}`);
      }
    });

  } catch (e) {
    console.error(`An unexpected error occurred: ${e.message}`);
  } finally {
    console.log("--- Script Execution Summary ---");
    console.log(`Processed Messages: ${processedMessages}`);
    console.log(`Downloaded Image Files: ${downloadedFiles}`);
    console.log(`Unzipped Files: ${unzippedFilesCount}`);
    console.log("Script execution completed.");
  }
}

/**
 * Gets or creates a folder in Google Drive.
 * If the folder already exists, it is returned. Otherwise, a new folder is created.
 *
 * @param {string} folderName The name of the folder.
 * @return {GoogleAppsScript.Drive.Folder} The Drive folder object.
 */
function getOrCreateFolder(folderName) {
  try {
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      console.log(`Folder '${folderName}' found in Google Drive.`);
      return folders.next();
    } else {
      console.log(`Folder '${folderName}' not found. Creating a new folder.`);
      return DriveApp.createFolder(folderName);
    }
  } catch (e) {
    console.error(`Failed to get or create folder '${folderName}'. Error: ${e.message}`);
    throw e;
  }
}

/**
 * Gets or creates a Gmail label.
 * If the label already exists, it is returned. Otherwise, a new label is created.
 *
 * @param {string} labelName The name of the label.
 * @return {GoogleAppsScript.Gmail.GmailLabel} The Gmail label object.
 */
function getOrCreateLabel(labelName) {
  try {
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      console.log(`Label '${labelName}' not found. Creating a new label.`);
      label = GmailApp.createLabel(labelName);
    } else {
      console.log(`Label '${labelName}' found.`);
    }
    return label;
  } catch (e) {
    console.error(`Failed to get or create label '${labelName}'. Error: ${e.message}`);
    throw e;
  }
}

/**
 * Checks if a file with the given name already exists in the specified folder.
 *
 * @param {GoogleAppsScript.Drive.Folder} folder The folder to search in.
 * @param {string} filename The name of the file to check for.
 * @return {boolean} True if the file exists, false otherwise.
 */
function fileExistsInFolder(folder, filename) {
  try {
    const files = folder.getFilesByName(filename);
    return files.hasNext();
  } catch (e) {
    console.error(`Failed to check for file '${filename}' in folder '${folder.getName()}'. Error: ${e.message}`);
    return false;
  }
}