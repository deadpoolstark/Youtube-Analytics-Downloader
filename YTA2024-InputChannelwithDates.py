from flask import Flask, render_template, request, send_file
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd
import re
import urllib.parse
import os

app = Flask(__name__)

# YouTube API configuration
DEVELOPER_KEY = ""  # Replace with your actual API key
YOUTUBE_API_SERVICE_NAME = "youtube"
YOUTUBE_API_VERSION = "v3"

def youtube_service():
    """Returns an instance of the YouTube API service."""
    return build(YOUTUBE_API_SERVICE_NAME, YOUTUBE_API_VERSION, developerKey=DEVELOPER_KEY)

def sanitize_filename(filename):
    """Sanitizes filenames to remove invalid characters and URL-encodes them."""
    filename = re.sub(r'[<>:"/\\|?*]', '', filename)  # Remove invalid characters
    filename = filename.replace(' ', '_')  # Replace spaces with underscores
    return urllib.parse.quote(filename)  # URL-encode the filename

def get_channel_id(youtube, channel_name):
    """Fetches channel ID by channel name."""
    search_response = youtube.search().list(
        q=channel_name,
        part="id",
        maxResults=1,
        type="channel"
    ).execute()

    if search_response['items']:
        return search_response['items'][0]['id']['channelId']
    return None

def get_channel_info(youtube, channel_id, channel_name):
    """Fetches channel statistics and metadata."""
    channels_response = youtube.channels().list(
        id=channel_id,
        part="statistics,snippet"
    ).execute()

    if channels_response['items']:
        channel = channels_response['items'][0]
        return {
            "Channel Name": channel_name,
            "Subscriber Count": channel['statistics'].get('subscriberCount', 'N/A'),
            "Total Views": channel['statistics'].get('viewCount', 'N/A'),
            "Joined Date": channel['snippet'].get('publishedAt', 'N/A').split('T')[0],
            "Location": channel['snippet'].get('country', 'N/A'),
            "Website": channel['snippet'].get('customUrl', 'N/A'),
            "Video Count": channel['statistics'].get('videoCount', 'N/A'),
            "Description": channel['snippet'].get('description', 'N/A')
        }
    return None

def get_all_videos(youtube, channel_id):
    """Fetches all videos for a channel, ensuring only video results are processed."""
    all_videos = []
    next_page_token = None

    while True:
        videos_response = youtube.search().list(
            channelId=channel_id,
            part="snippet",
            order="date",
            maxResults=50,  # Fetch 50 videos per page (maximum allowed by YouTube API)
            pageToken=next_page_token
        ).execute()

        for item in videos_response.get('items', []):
            # Check if the result is a video by verifying if 'videoId' exists
            if item['id'].get('videoId'):
                all_videos.append(item)

        next_page_token = videos_response.get('nextPageToken', None)

        # Break the loop if no more pages are available
        if not next_page_token:
            break

    return all_videos

def get_video_data(youtube, video_id):
    """Fetches detailed video statistics."""
    response = youtube.videos().list(
        part="statistics,snippet,contentDetails,status,localizations",
        id=video_id
    ).execute()

    if response['items']:
        video = response['items'][0]
        return {
            "View Count": video['statistics'].get('viewCount', 'N/A'),
            "Like Count": video['statistics'].get('likeCount', 'N/A'),
            "Description": video['snippet'].get('description', 'N/A'),
            "Duration": video['contentDetails'].get('duration', 'N/A'),
            "Published At": video['snippet'].get('publishedAt', 'N/A'),
            "Subtitles Available": video['status'].get('caption', 'N/A') == 'true',
            "Subtitle Language": ', '.join([
                caption['snippet']['language']
                for caption in youtube.captions().list(part='snippet', videoId=video_id).execute().get('items', [])
            ]) or 'N/A',
            "Default Audio Language": video['snippet'].get('defaultAudioLanguage', 'N/A'),
            "Default Language": video['snippet'].get('defaultLanguage', 'N/A'),
            "Localized Titles": ', '.join([
                f"{lang}: {localization['title']}"
                for lang, localization in video.get('localizations', {}).items()
            ]) or 'N/A'
        }
    return None

def get_channel_and_video_data(channel_name):
    """Fetches channel and video data, and saves to an Excel file."""
    youtube = youtube_service()
    channel_id = get_channel_id(youtube, channel_name)
    
    if not channel_id:
        return None

    # Fetch channel info
    channel_info = get_channel_info(youtube, channel_id, channel_name)

    # Fetch all videos for the channel
    videos = get_all_videos(youtube, channel_id)
    video_data = []
    for idx, video in enumerate(videos, start=1):
        video_id = video['id']['videoId']
        video_info = get_video_data(youtube, video_id)
        video_info.update({
            "Index": idx,
            "Title": video['snippet']['title'],
            "Link": f"https://www.youtube.com/watch?v={video_id}"
        })
        video_data.append(video_info)

    # Save to Excel
    output_file = f"{sanitize_filename(channel_name)}.xlsx"
    channel_df = pd.DataFrame([channel_info])
    video_df = pd.DataFrame(video_data)
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        channel_df.to_excel(writer, sheet_name='Channel Info', index=False)
        video_df.to_excel(writer, sheet_name='Videos', index=False)

    return output_file

@app.route('/', methods=['GET', 'POST'])
def index():
    """Handles the main page with form submission."""
    if request.method == 'POST':
        channel_name = request.form['channel_name']
        result = get_channel_and_video_data(channel_name)

        if result:
            return send_file(result, as_attachment=True)
        else:
            return render_template('index.html', error_message="Channel not found or an error occurred.")

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
