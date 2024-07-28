import json
import datetime
import argparse
from docx import Document

def transcribe(filename):
    # Take the input as the filename
    filename_prefix = filename.split('.')[0]

    # Create a Document object
    doc = Document()

    with open(filename) as f:
        data = json.load(f)
        speaker_start_times = {}

        # Process speaker labels and items simultaneously
        for segment in data['results']['speaker_labels']['segments']:
            for item in segment['items']:
                speaker_start_times[item['start_time']] = item['speaker_label']

        lines = []
        line = ''
        current_speaker = None
        start_time = None

        # Process items
        for item in data['results']['items']:
            if 'start_time' in item:
                current_speaker = speaker_start_times.get(item['start_time'], current_speaker)
                if current_speaker != speaker_start_times.get(start_time):
                    if start_time is not None:
                        lines.append({'speaker': current_speaker, 'line': line, 'time': start_time})
                    line = item['alternatives'][0]['content']
                    start_time = item['start_time']
                else:
                    line += ' ' + item['alternatives'][0]['content']
            elif item['type'] == 'punctuation':
                line += item['alternatives'][0]['content']

        lines.append({'speaker': current_speaker, 'line': line, 'time': start_time})

        # Filter out None values
        lines = [line for line in lines if line['time'] is not None]

        # Sort the results by time
        sorted_lines = sorted(lines, key=lambda k: float(k['time']))

        # Add lines to the document
        for line_data in sorted_lines:
            time = datetime.timedelta(seconds=int(round(float(line_data['time']))))
            line = f"[{time}] {line_data['speaker']}: {line_data['line']}"
            doc.add_paragraph(line)

    # Save the document
    doc.save(f"{filename_prefix}_result.docx")

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument("-f", "--file", dest="filename",
                    help="file to be transcribed", metavar="FILE")
    
    args = parser.parse_args()
    
    transcribe(args.filename)
