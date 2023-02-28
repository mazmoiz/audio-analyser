from emotion_recognition import EmotionRecognizer
import os
import transcribeUtils as tu
import audioUtils as au
import utils as u
import json
from pathlib import Path
import excel as e

# IMPORTANT - Get the API key from https://www.assemblyai.com/ 

def main():
    
    # Create header with authorization along with content-type
    header = {
        'authorization': '<YOUR API KEY>',
        'content-type': 'application/json',
        'Accept': 'application/json'
    }

    # Clean the audio_chunks folder.
    u.clean_output_folder('audio_chunks')

    # Specify your audio path here...
    audio_file_path = "2_people_audio1.mp3"

    # Get the upload URL for the audio file.
    upload_url = tu.upload_file(audio_file_path, header)

    # Request a transcription
    transcript_response = tu.request_transcript(upload_url, header)

    # Create a polling endpoint that will let us check when the transcription is complete
    polling_endpoint = tu.make_polling_endpoint(transcript_response)

    # Wait until the transcription is complete
    tu.wait_for_completion(polling_endpoint, header)

    # Request the paragraphs of the transcript
    speakers = tu.get_response(polling_endpoint, header)
    s = json.dumps(speakers)
    u.write_file(s, 'speakers.json')


    # Get the file extension
    file_extension = os.path.splitext(audio_file_path)[1]
    wav_fpath = audio_file_path

    # Convert the entire mp3 file to wav
    if file_extension == '.mp3':
        wav_fpath = au.mp3_to_wav(audio_file_path)


    # Load the JSON from file.
    with open('speakers.json') as speakers_file:
        speaker_json = json.load(speakers_file)
    

    estimators = u.get_best_estimators(True)
    estimators_str, estimator_dict = u.get_estimators_name(estimators)
    features = ["mfcc", "chroma", "mel"]

    output = []
    detector = EmotionRecognizer(estimator_dict["BaggingClassifier"], emotions="neutral,calm,happy,sad,angry,fear,disgust,ps,boredom".split(","), features=features, verbose=0)

    # Loop through the speaker list and create audio chunks.
    for x, speaker in enumerate(speaker_json):
        chunk_audio_path = au.create_audio_chunks(wav_fpath, speaker['start'], speaker['end'], speaker['speaker'], x)
        emotion = detector.predict(chunk_audio_path)
        #print(str(x), ' - ' , speaker['speaker'], ' - ', speaker['text'], '(' + emotion + ')')
        #out = str(x) + ' - ' + speaker['speaker'] + ' - ' + speaker['text'] + ' - ' + emotion
        #emotions.append(out)
        dbfs = au.get_dbfs(chunk_audio_path, speaker['start'], speaker['end'])
        #print('DFBS -> ', dbfs)
        fo = {
                "speaker": speaker['speaker'],
                "text": speaker['text'],
                "start": speaker['start'],
                "end": speaker['end'],
                "sentiment": speaker['sentiment'],
                "emotion": emotion,
                "dbfs": dbfs
            }
        output.append(fo)

    parsed_output = json.dumps(output)
    u.write_file(parsed_output, str(Path.cwd()) + '/report/src/final_output.json')
    print(parsed_output)

    e.create_excel(str(Path.cwd()) + '/report/src/final_output.json', os.path.basename(audio_file_path) + '__' + u.get_localtime())

if __name__ == '__main__':
    main()
