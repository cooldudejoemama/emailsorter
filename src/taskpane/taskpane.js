Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
      document.getElementById("organize-btn").onclick = organizeEmail;
      displayFolders();
  }
});

// Global variables
const API_ENDPOINT = 'http://localhost:3000/classify';  // During development
const FOLDERS = ['Personal', 'Work', 'Finance', 'Shopping', 'Social', 'Other'];

async function organizeEmail() {
  try {
      setStatus('Analyzing email...');
      
      // Get current email
      const item = Office.context.mailbox.item;
      
      // Get email content
      const subject = item.subject;
      const body = await new Promise((resolve) => {
          item.body.getAsync('text', {}, (result) => {
              resolve(result.value);
          });
      });

      // Call our backend API
      const response = await fetch(API_ENDPOINT, {
          method: 'POST',
          headers: {
              'Content-Type': 'application/json',
          },
          body: JSON.stringify({
              subject: subject,
              body: body
          })
      });

      if (!response.ok) {
          throw new Error('API request failed');
      }

      const data = await response.json();
      const suggestedFolder = data.folder;

      // Move email to suggested folder
      await moveToFolder(suggestedFolder);
      
      setStatus(`Email moved to ${suggestedFolder}`);

  } catch (error) {
      setStatus('Error: ' + error.message);
      console.error('Error:', error);
  }
}

async function moveToFolder(folderName) {
  return new Promise((resolve, reject) => {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async (result) => {
          if (result.status === "succeeded") {
              const accessToken = result.value;
              const item = Office.context.mailbox.item;
              
              try {
                  // Get the REST ID of the message
                  const restId = Office.context.mailbox.convertToRestId(
                      item.itemId,
                      Office.MailboxEnums.RestVersion.v2_0
                  );

                  // Get user's email
                  const userEmail = Office.context.mailbox.userProfile.emailAddress;
                  
                  // Microsoft Graph API endpoint
                  const graphEndpoint = `https://graph.microsoft.com/v2.0/users/${userEmail}/messages/${restId}/move`;
                  
                  // Create or get the destination folder ID
                  const folderId = await getFolderId(folderName, accessToken);
                  
                  // Move the email
                  const response = await fetch(graphEndpoint, {
                      method: 'POST',
                      headers: {
                          'Authorization': `Bearer ${accessToken}`,
                          'Content-Type': 'application/json'
                      },
                      body: JSON.stringify({
                          destinationId: folderId
                      })
                  });

                  if (!response.ok) {
                      throw new Error('Failed to move email');
                  }

                  resolve();
                  
              } catch (error) {
                  reject(error);
              }
          } else {
              reject(new Error('Could not get access token'));
          }
      });
  });
}

async function getFolderId(folderName, accessToken) {
  const userEmail = Office.context.mailbox.userProfile.emailAddress;
  const endpoint = `https://graph.microsoft.com/v2.0/users/${userEmail}/mailFolders`;
  
  // First, try to find the existing folder
  const response = await fetch(endpoint, {
      headers: {
          'Authorization': `Bearer ${accessToken}`
      }
  });

  if (!response.ok) {
      throw new Error('Failed to get folders');
  }

  const data = await response.json();
  const folder = data.value.find(f => f.displayName === folderName);

  if (folder) {
      return folder.id;
  }

  // If folder doesn't exist, create it
  const createResponse = await fetch(endpoint, {
      method: 'POST',
      headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
      },
      body: JSON.stringify({
          displayName: folderName
      })
  });

  if (!createResponse.ok) {
      throw new Error('Failed to create folder');
  }

  const newFolder = await createResponse.json();
  return newFolder.id;
}

function displayFolders() {
  const folderList = document.getElementById('folder-list');
  folderList.innerHTML = '<h3>Available Folders:</h3>';
  
  const ul = document.createElement('ul');
  ul.className = 'ms-List';
  
  FOLDERS.forEach(folder => {
      const li = document.createElement('li');
      li.className = 'ms-ListItem';
      li.textContent = folder;
      ul.appendChild(li);
  });
  
  folderList.appendChild(ul);
}

function setStatus(message) {
  const status = document.getElementById('status');
  status.textContent = message;
}