import wave
from pathlib import Path
from pydub import AudioSegment, utils

frame_rate = 16000

def mp3_to_wav(audio_file_path):
    sound = AudioSegment.from_mp3(audio_file_path)
    audio_file_path = audio_file_path.split('.')[0] + '.wav'
    sound = sound.set_frame_rate(frame_rate)
    sound = sound.set_channels(1)
    sound.export(audio_file_path, format="wav")
    return audio_file_path

def create_audio_chunks(wav_fpath, start, end, name, marker):
    with wave.open(wav_fpath, "rb") as infile:
        # get file data
            nchannels = infile.getnchannels()
            sampwidth = infile.getsampwidth()
            framerate = infile.getframerate()
            # set position in wave to start of segment
            infile.setpos(int(start/1000 * framerate))
            # extract data
            data = infile.readframes(
                int(((end/1000) - (start/1000)) * framerate))

    output_to_file = str(Path.cwd()) + '/audio_chunks/' + \
        str(marker) + '_chunk_' + str(start/1000) + '-' + \
        str(end/1000) + '_' + str(name) + '.wav'
    
    with wave.open(output_to_file, 'w') as outfile:
            outfile.setnchannels(nchannels)
            outfile.setsampwidth(sampwidth)
            outfile.setframerate(framerate)
            outfile.setnframes(int(len(data) / sampwidth))
            outfile.writeframes(data)
    
    return output_to_file

chunk_length = 100 #100 ms
def get_dbfs(audio_chunk_file_path, start, end):
    audio = AudioSegment.from_wav(audio_chunk_file_path)
    audio_chunks = utils.make_chunks(audio, chunk_length) 
    
    dbsf = []

    for x, chunk in enumerate(audio_chunks):
        if x == len(audio_chunks)-1:
            start = end
        else:
            start = start + chunk_length

        d = {
            "time": start,
            "value": chunk.dBFS
        }
        
        dbsf.append(d)

    return dbsf