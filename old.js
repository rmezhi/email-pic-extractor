

function downloadPolarBearPhotos() {
  const SENDER_EMAIL = 'no-reply@kaymbu.com'; // Replace with the actual sender's email
  const SEARCH_QUERY = `"Polar Bears (Twos)"`;
  const FOLDER_NAME = 'daycare-pics';
  
  let folder = getOrCreateFolder(FOLDER_NAME);
  
  const threads = GmailApp.search(SEARCH_QUERY);
  
  if (threads.length === 0) {
    Logger.log('No matching emails found.');
    return;
  }
  
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      try {
        const body = message.getPlainBody();
        const urlMatch = body.match(/https:\/\/email3\.kaymbu\.com\/ls\/click\?upn=[^\s]+/);
        
        if (urlMatch) {
          const trackingUrl = urlMatch[0].replace(/=\r?\n/g, ''); // Remove soft line breaks
          
          // Use UrlFetchApp to get the redirect URL
          const response = UrlFetchApp.fetch(trackingUrl, {
            muteHttpExceptions: true,
            followRedirects: false // This is the key change
          });
          
          const finalImageUrl = response.getHeaders().Location;
          
          if (finalImageUrl) {
            // Fetch the image from the final URL
            const imageResponse = UrlFetchApp.fetch(finalImageUrl, { muteHttpExceptions: true });
            
            if (imageResponse.getResponseCode() === 200) {
              const blob = imageResponse.getBlob();
              const filename = `PolarBear_${new Date().toISOString()}.jpg`;
              folder.createFile(blob.setName(filename));
              
              Logger.log(`Successfully downloaded: ${filename}`);
              message.markRead();
              
            } else {
              Logger.log(`Failed to fetch image from final URL: ${finalImageUrl}. Response code: ${imageResponse.getResponseCode()}`);
            }
          } else {
            Logger.log(`No redirect URL found in headers for: ${trackingUrl}`);
          }
        }
      } catch (e) {
        Logger.log(`Error processing message from ${message.getFrom()}: ${e.message}`);
      }
    });
  });
}

/**
 * Gets or creates a folder in Google Drive.
 * @param {string} folderName The name of the folder.
 * @return {GoogleAppsScript.Drive.Folder} The Drive folder object.
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
    }
}