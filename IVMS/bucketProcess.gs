function uploadImageToGCS(id) {
  var bucketName = "testing-fonnte-bucket"; // Replace with your GCS bucket name
  var fileBlob = DriveApp.getFileById(id).getBlob(); // Get file from Google Drive

  // Dynamically get Content-Type and filename from the Blob
  var contentType = fileBlob.getContentType();
  var originalFilename = fileBlob.getName();
  var fileExtension = originalFilename.substring(originalFilename.lastIndexOf('.')); // Extract extension (e.g., ".jpg", ".png")
  var fileName = "uploaded-image-" + Utilities.getUuid() + fileExtension; // Create a unique filename with original extension

  // Retrieve the service account key
  var serviceAccountKeyJson = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY');
  if (!serviceAccountKeyJson) {
    Logger.log("Service account key not found in Script Properties.");
    return null;
  }
  var serviceAccountKey = JSON.parse(serviceAccountKeyJson);

  // Get the access token
  var accessToken = getAccessToken(serviceAccountKey);
  if (!accessToken) {
    Logger.log("Failed to retrieve access token.");
    return null;
  }

  var url = "https://storage.googleapis.com/upload/storage/v1/b/" + bucketName + "/o?uploadType=media&name=" + fileName;

  var options = {
    method: "post",
    contentType: contentType, // Use the dynamic Content-Type!
    payload: fileBlob.getBytes(),
    headers: {
      "Authorization": "Bearer " + accessToken,
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());

  if (response.getResponseCode() == 200) {
    Logger.log("Upload successful! Image URL: https://storage.googleapis.com/" + bucketName + "/" + fileName + ", Content-Type: " + contentType);
    return "https://storage.googleapis.com/" + bucketName + "/" + fileName;
  } else {
    Logger.log("Upload failed: " + response.getContentText());
    return null;
  }
}

function deleteImageFromGCS(fileName) {
  const bucketName = "testing-fonnte-bucket";
  Logger.log("Attempting to delete file: " + fileName + " from bucket: " + bucketName);

  if (!bucketName || !fileName) {
    Logger.log("Error: Bucket name and file name are required for deletion.");
    return false;
  }

  // Retrieve the service account key
  var serviceAccountKeyJson = PropertiesService.getScriptProperties().getProperty('SERVICE_ACCOUNT_KEY');
  if (!serviceAccountKeyJson) {
    Logger.log("Service account key not found in Script Properties.");
    return false;
  }
  var serviceAccountKey = JSON.parse(serviceAccountKeyJson);

  // Get the access token
  var accessToken = getAccessToken(serviceAccountKey);
  if (!accessToken) {
    Logger.log("Failed to retrieve access token for deletion.");
    return false;
  }

  // Construct the DELETE request URL
  var url = "https://storage.googleapis.com/storage/v1/b/" + bucketName + "/o/" + encodeURIComponent(fileName);

  var options = {
    method: "delete", // Set method to DELETE
    headers: {
      "Authorization": "Bearer " + accessToken,
    },
    muteHttpExceptions: true // To get detailed error responses
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    if (responseCode == 204) { // 204 No Content is successful for DELETE
      Logger.log("Deletion successful! File: " + fileName + " deleted from bucket: " + bucketName);
      return true;
    } else {
      Logger.log("Deletion failed. Response Code: " + responseCode);
      Logger.log("Response Content: " + response.getContentText());
      return false;
    }
  } catch (e) {
    Logger.log("Error during deletion: " + e);
    return false;
  }
}

function getAccessToken(serviceAccountKey) {
  var jwtToken = getJwtToken(serviceAccountKey);

  var tokenResponse = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwtToken
    },
    muteHttpExceptions: true
  });

  var responseCode = tokenResponse.getResponseCode();
  if (responseCode === 200) {
    var tokenData = JSON.parse(tokenResponse.getContentText());
    return tokenData.access_token; // Return the access token
  } else {
    Logger.log("Failed to retrieve access token: " + tokenResponse.getContentText());
    return null;
  }
}


function getJwtToken(serviceAccountKey) {
  var jwtHeader = Utilities.base64Encode(JSON.stringify({ alg: "RS256", typ: "JWT" }));
  var jwtClaimSet = Utilities.base64Encode(JSON.stringify({
    iss: serviceAccountKey.client_email,
    scope: "https://www.googleapis.com/auth/devstorage.read_write", // Correct scope for GCS
    aud: "https://oauth2.googleapis.com/token",
    exp: Math.floor(new Date().getTime() / 1000) + 3600, // Token expires in 1 hour
    iat: Math.floor(new Date().getTime() / 1000)
  }));

  var signatureInput = jwtHeader + "." + jwtClaimSet;
  var signature = Utilities.computeRsaSha256Signature(signatureInput, serviceAccountKey.private_key);
  var jwtSignature = Utilities.base64Encode(signature);

  return signatureInput + "." + jwtSignature;
}
