import streamlit as st
import docx2txt
from io import BytesIO
import base64
import docx
from audio_recorder_streamlit import audio_recorder
from google.cloud import speech_v1p1beta1 as speech
from google.cloud import texttospeech
import os
from pydub import AudioSegment
from io import BytesIO

st.set_page_config(page_title="Question/Answer App", layout="wide")

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'key.json'

def transcribe_audio(mono_audio_bytes):
    client = speech.SpeechClient()
    audio = speech.RecognitionAudio(content=mono_audio_bytes)
    config = speech.RecognitionConfig(
        encoding=speech.RecognitionConfig.AudioEncoding.LINEAR16,
        sample_rate_hertz=48000,
        language_code="en-US",
        model="default"
    )

    response = client.recognize(config=config, audio=audio)
    transcript = ""

    for result in response.results:
        transcript += result.alternatives[0].transcript

    return transcript


def synthesize_speech(text):
    client = texttospeech.TextToSpeechClient()
    input_text = texttospeech.SynthesisInput(text=text)
    voice = texttospeech.VoiceSelectionParams(
        language_code="en-US", ssml_gender=texttospeech.SsmlVoiceGender.FEMALE
    )
    audio_config = texttospeech.AudioConfig(
        audio_encoding=texttospeech.AudioEncoding.LINEAR16
    )

    response = client.synthesize_speech(
        input=input_text, voice=voice, audio_config=audio_config
    )

    return response.audio_content

def audio():
    st.title("Question/Answer App")
    st.write("This app helps you to extract questions from a Word document, record your answers, and save them back into the document.")

    if "answers" not in st.session_state:
        st.session_state.answers = {}

    uploaded_file = st.file_uploader("Choose a Word document", type="docx")

    if uploaded_file is not None:
        file_contents = uploaded_file.read()
        text = docx2txt.process(BytesIO(file_contents))

        questions = []
        for line in text.split("\n"):
            line = line.strip()
            if line.endswith("?"):
                questions.append(line)

        question_index = st.number_input("Current question index:", min_value=0, max_value=len(questions) - 1, value=0, step=1)

        st.write("### Extracted questions:")
        st.write(questions[question_index])

        if st.button("Speak current question"):
            synthesized_audio = synthesize_speech(questions[question_index])
            st.audio(synthesized_audio, format="audio/wav")

        recorder = audio_recorder()
        audio_bytes = recorder

        if audio_bytes:
            st.audio(audio_bytes, format="audio/wav")

            stereo_audio = AudioSegment.from_wav(BytesIO(audio_bytes))
            mono_audio = stereo_audio.set_channels(1).set_frame_rate(48000)
            mono_audio_bytes = mono_audio.export(format="wav").read()

            if st.button("Convert to text"):
                answer = transcribe_audio(mono_audio_bytes)
                st.write("### Answer:")
                st.write(answer)

                st.session_state.answers[question_index] = answer

        if st.button("Save answers to the document"):
            doc = docx.Document(BytesIO(file_contents))

            title = doc.add_paragraph("Question and Answer")
            title.runs[0].font.size = docx.shared.Pt(24)
            title.runs[0].font.bold = True
            doc.add_paragraph("")

            for index, answer in st.session_state.answers.items():
                question_paragraph = doc.add_paragraph()
                question_paragraph.add_run("Question {}: ".format(index + 1)).bold = True
                question_paragraph.add_run(questions[index])
                answer_paragraph = doc.add_paragraph()
                answer_paragraph.add_run("Answer {}: ".format(index + 1)).bold = True
                answer_paragraph.add_run(answer)

            try:
                output_filename = "output_" + uploaded_file.name
                with BytesIO() as output_file:
                    doc.save(output_file)
                    output_file.seek(0)
                    content = output_file.read()
                st.success("Answers saved to the '{}'.".format(output_filename))

                # Add the download button
                st.download_button(
                    label="Download updated document",
                    data=content,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except PermissionError:
                st.error("Permission denied: Please close the current '{}' file before saving.".format(uploaded_file.name))

    else:
        st.error("Please upload a document before attempting to save answers.")
        


if __name__ == "__main__":
    audio()
