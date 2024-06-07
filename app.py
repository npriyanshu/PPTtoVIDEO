from flask import Flask, request, jsonify, render_template, url_for
from pptx import Presentation
from gtts import gTTS
from flask_cors import CORS
from moviepy.editor import concatenate_videoclips, ImageClip, AudioFileClip
from PIL import Image
from io import BytesIO
import os
import comtypes.client
import pythoncom

app = Flask(__name__)
CORS(app, origins='*', allow_headers=['Content-Type'])

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_ppt_to_video():
    file = request.files['file']
    if not file:
        return jsonify({"error": "No file provided"}), 400

    # Create a directory to store intermediate files
    media_dir = os.path.abspath('media')
    if not os.path.exists(media_dir):
        os.makedirs(media_dir)

    # Load the PowerPoint presentation
    ppt_file_path = os.path.join(media_dir, file.filename)
    file.save(ppt_file_path)
    prs = Presentation(ppt_file_path)

    # Convert slides to images and generate audio for each slide
    image_files = []
    audio_files = []
    for i, slide in enumerate(prs.slides):
        # Save slide as image
        slide_image_path = os.path.join(media_dir, f"slide_{i}.png")
        save_slide_as_image(ppt_file_path, i, slide_image_path)
        image_files.append(slide_image_path)

        # Extract notes text
        notes_text = ''
        if slide.notes_slide and slide.notes_slide.notes_text_frame:
            notes_text = slide.notes_slide.notes_text_frame.text.strip()

        # Convert notes text to speech and save as audio file
        if notes_text:
            audio_file_path = os.path.join(media_dir, f"audio_{i}.mp3")
            tts = gTTS(text=notes_text, lang='en')
            tts.save(audio_file_path)
            audio_files.append(audio_file_path)
        else:
            audio_files.append(None)  # Append None if there's no audio

    # Create video clips for each slide
    video_clips = []
    for img_file, audio_file in zip(image_files, audio_files):
        img_clip = ImageClip(img_file)
        if audio_file:
            audio_clip = AudioFileClip(audio_file)
            img_clip = img_clip.set_duration(audio_clip.duration)
            video_clip = img_clip.set_audio(audio_clip)
        else:
            img_clip = img_clip.set_duration(2)  # Default duration for silent slides
            video_clip = img_clip
        video_clips.append(video_clip)

    # Concatenate video clips
    final_video = concatenate_videoclips(video_clips, method="compose")
    output_video_path = os.path.join('static', 'output_video.mp4')
    final_video.write_videofile(output_video_path, fps=24)

    # Clean up intermediate files
    for img in image_files:
        os.remove(img)
    for audio in audio_files:
        if audio:
            os.remove(audio)

    # Remove the uploaded ppt file
    os.remove(ppt_file_path)

    return jsonify({"video_url": url_for('static', filename='output_video.mp4')})

def save_slide_as_image(ppt_file_path, slide_index, output_image_path):
    # Initialize COM library
    pythoncom.CoInitialize()
    try:
        # Use comtypes to export slide as image
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(os.path.abspath(ppt_file_path))
        slide = presentation.Slides[slide_index + 1]  # Slide index is 1-based in PowerPoint
        slide.Export(os.path.abspath(output_image_path), "PNG")
        presentation.Close()
    finally:
        powerpoint.Quit()
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    app.run(debug=True)
