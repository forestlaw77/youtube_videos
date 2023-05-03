/*
 * Get a list of Youtube videos and output to Google Spread Sheet. 
 * 
 * Tsutomu Funada
 */

/*
 * Format date in user timezone
 */
function formatDateInUserTimeZone(timestamp) {
  const userTimeZone = Session.getScriptTimeZone();
  const format = "yyyy/MM/dd HH:mm:ss";
  const dateFormat = Utilities.formatDate(new Date(timestamp), userTimeZone, format);
  return dateFormat;
}

/*
 * Format Duration
 * 
 * e.g.) PT4MS5S -> 4:05
 */
function formatDuration(duration) {
  const match = duration.match(/PT(\d+H)?(\d+M)?(\d+S)?/);
  const hours = (parseInt(match[1]) || 0);
  const minutes = (parseInt(match[2]) || 0);
  const seconds = (parseInt(match[3]) || 0);
  return `${hours}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
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
 * Get All Video Infomation
 */
function getAllVideoInfo (channelId) {
  let videos = [];
  let nextPageToken = "";

  do {
    let searchResult = {};
    let videoResult = {};
    
    try {
      searchResult = YouTube.Search.list("id", {
        channelId: channelId,
        type: "video",
        maxResults: 50,
        pageToken: nextPageToken,
      });
      const videoIds = searchResult.items.map(item => item.id.videoId).join(",");
      videoResult = YouTube.Videos.list("snippet, contentDetails, statistics", {
        id: videoIds,
      });
    } catch (error) {
      console.error("Error fetching video or video Info: ", error);
      continue;
    }

    for (let videoIndex = 0; videoIndex < searchResult.items.length; videoIndex++) {
      const video = searchResult.items[videoIndex];
      const videoInfo = videoResult.items[videoIndex];
      videos.push({
        videoId: video.id.videoId,
        channelId: channelId,
        channelTitle: videoInfo.snippet.channelTitle,
        publishedAt: videoInfo.snippet.publishedAt,
        title: videoInfo.snippet.title,
        description: videoInfo.snippet.description,
        duration: videoInfo.contentDetails.duration,
        viewCount: videoInfo.statistics.viewCount,
        likeCount: videoInfo.statistics.likeCount,
      });
    }

    nextPageToken = searchResult.nextPageToken;
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
    return null;
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
    return null;
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
 * Write the video info to the spreadsheet
 * 
 * Code optimization: The old writeVideoInfoToSpreadsheet function uses a loop to append each video row to the spreadsheet.
 * However, this can be slow when dealing with a large number of rows.
 * A better approach is to write the video data to a 2D array and then use the setValues method to write the entire array to the sheet at once.
 * This can significantly improve performance.
 * 
 */
function writeVideoInfoToSpreadsheet(sheet, videos) {

  // Write header
  writeHeader(sheet);

  // Create an empty 2D array to hold the video data
  const data = [];
  
  // Loop through the videos and add earch video row to the 2D array
  videos.forEach(video => {
    const videoRow = [
      video.videoId,
      video.channelId,
      video.channelTitle,
      formatDateInUserTimeZone(video.publishedAt), // Changed to format according to the user's time zone.
      '=HYPERLINK("https://www.youtube.com/watch?v=' + video.videoId + '", "' + video.title + '")',
      video.description,
      formatDuration(video.duration), 
      video.viewCount,
      video.likeCount,
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
  if (!handleId) {
    throw new Error("No Handle ID found. Please set handle ID to the sheet name.");
  }

  // Get channel ID
  const channelId = getChannelId(handleId);
  if (!channelId) {
    throw new Error("Channel ID is not available. Please check handle ID.");
  }

  // Get all video info
  const videos = getAllVideoInfo(channelId);

  // Write the all video info to spreadsheet
  writeVideoInfoToSpreadsheet(sheet, videos);

}

