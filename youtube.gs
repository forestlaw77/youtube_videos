/*
 * Get a list of Youtube videos and output to Google Spread Sheet. 
 * 
 * Tsutomu Funada
 */

/*
 * Convert Japan Time
 */
function dateJST(timestamp) {
  return Utilities.formatDate(new Date(timestamp), 'JST', 'yyyy/MM/dd HH:mm:ss');
}

/*
 * Trigger runs automatically when a user opens a spreadsheet.
 */
function onOpen() {
  // Add dedicated menu to UI
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('MyYouTube');
  menu.addItem('Get video information', 'main');
  menu.addToUi();
}

/*
 * Get All Videos 
 */
function getAllVideos(channelId) {
  let videos = [];
  let nextPageToken = "";

  do {
    let response = [];
    try {
      response = YouTube.Search.list("id, snippet", {
        channelId,
        type: "video",
        maxResults: 50,
        pageToken: nextPageToken,
      });
    } catch (error) {
      console.error("Error fetching videos: ", error);
      break;
    }
    videos = videos.concat(response.items);
    nextPageToken = response.nextPageToken;
  } while (nextPageToken);

  return videos;
}

/*
 * Get channel ID
 */
function getChannelId(handleId) {
  // Get the channel ID from the YouTube API
  let response;
  try {
    response = YouTube.Search.list("id", {
      type: "channel",
      q: handleId,
    });
  } catch (error) {
    console.error("Error fetching channelId: ", error);
  }
  const channelId = response.items[0].id.channelId;
  return channelId;
}

/*
 * Get handle ID
 */
function getHandleId(sheet) {
  const handleId = sheet.getSheetName();
  if (!handleId.match(/^@.*/)) { // Is valid handle ID?
    throw new Error("No Handle ID found. Please set handle ID to sheet name.");
  }
  return handleId;
}

/*
 * Write a header
 */
function writeHeader(sheet) {
  // Get the header from the active sheet.
  const headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()[0];
  // Clear the sheet & re-write the header.
  sheet.clear();
  sheet.getRange(1,1,1,headers.length).setValues([headers]);
}

/*
 * write a footer
 */
function writeFooter(sheet) {
  sheet.getRange(sheet.getLastRow()+2,1,1,1).setValue("done");
}

/*
 * Write the videos to a spreadsheet
 * 
 * Code optimization: The old writeVideosToSpreadsheet function uses a loop to append each video row to the spreadsheet.
 * However, this can be slow when dealing with a large number of rows.
 * A better approach is to write the video data to a 2D array and then use the setValues method to write the entire array to the sheet at once.
 * This can significantly improve performance.
 * 
 */
function writeVideosToSpreadsheet(sheet, videos) {

  // Write header
  writeHeader(sheet);

  // Create an empty 2D array to hold the video data
  const data = [];
  
  // Loop through the videos and add earch video row to the 2D array
  videos.forEach(video => {
    const videoRow = [
      video.snippet.channelTitle,
      video.snippet.publishedAt,
      '=HYPERLINK("https://www.youtube.com/watch?v=' + video.id.videoId + '", "' + video.snippet.title + '")',
      video.snippet.description,
    ];
    data.push(videoRow);
  });

  sheet.getRange(sheet.getLastRow()+1, 1, data.length, data[0].length).setValues(data);

  // Write footer
  writeFooter(sheet);
}

/*
 * Main function
 */
function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Get hanndle ID
  const handleId = getHandleId(sheet);

  // Get channel ID
  const channelId = getChannelId(handleId);

  // Get videos
  const videos = getAllVideos(channelId);

  // Write videos to spreadsheet
  writeVideosToSpreadsheet(sheet, videos);
}

