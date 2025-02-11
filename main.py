import pandas as pd
from pytube import YouTube
from moviepy.editor import VideoFileClip
import os


# Function to download the video from a given URL
def download_video(video_id, youtube_url, output_directory):
    try:
        yt = YouTube(youtube_url)
        video_stream = yt.streams.filter(file_extension="mp4").first()
        video_output_filename = f"{video_id}.mp4"
        video_output_path = os.path.join(output_directory, video_output_filename)
        video_stream.download(output_directory, filename=video_output_filename)
        return video_output_path, None  # Return None for error if download is successful
    except Exception as e:
        return None, str(e)  # Return the error message if there's an exception


# Function to convert video to audio
def convert_video_to_audio(video_filename, audio_output_filename):
    video_clip = VideoFileClip(video_filename)
    audio_clip = video_clip.audio
    audio_output_path = os.path.join(os.path.dirname(video_filename), audio_output_filename)
    audio_clip.write_audiofile(audio_output_path)
    video_clip.close()
    audio_clip.close()
    return audio_output_path


# Load the Excel file
excel_file_path = "C:/Users/ravin/OneDrive/Documents/YOUTUBEDATA.xlsx"
df = pd.read_excel(excel_file_path)

# Set the output directory
output_directory = "C:/Users/ravin/VSCODE/MODELFILES"

# Find the next available video to download
for index, row in df.iterrows():
    if row['Video Status'] != 'downloaded':
        video_id = row['Video ID']
        youtube_url = row['Video URL']
        video_output_path, error = download_video(video_id, youtube_url, output_directory)

        if error:
            print(f"Error downloading video {video_id}: {error}")
            continue  # Move to the next video if there's an error

        # Convert video to audio
        audio_output_filename = f"{video_id}.mp3"
        audio_output_path = convert_video_to_audio(video_output_path, audio_output_filename)

        # Update the Excel file with the download status
        df.at[index, 'Video Status'] = 'downloaded'
        df.to_excel(excel_file_path, index=False)

        print(f"Video {video_id} downloaded to: {video_output_path}")
        print(f"Audio {audio_output_filename} converted to: {audio_output_path}")
        print("Excel file updated.")
        break  # Exit the loop after downloading the first available video
else:
    print("All videos have already been downloaded.")
