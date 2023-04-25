import streamlit as st
import docx2txt
import pyttsx3
from io import BytesIO
import speech_recognition as sr
import docx
import base64
from audio_recorder_streamlit import audio_recorder
st.set_page_config(page_title="Question/Answer App", layout="wide")

def get_docx_download_link(filename, content):
    b64 = base64.b64encode(content).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="{filename}">Download {filename}</a>'

def stop_engine(engine):
    engine.stop()

def audio():
    st.title("Question/Answer App")
    st.write("This app helps you to extract questions from a Word document, record your answers, and save them back into the document.")
    
    # Initialize session_state
    if "answers" not in st.session_state:
        st.session_state.answers = {}

    # Define the file uploader
    uploaded_file = st.file_uploader("Choose a Word document", type="docx")

    # Check if a file was uploaded
    if uploaded_file is not None:

        st.info("A document has been uploaded. Please follow the steps below.")

        # Use docx2txt to extract text from the Word document
        file_contents = uploaded_file.read()
        text = docx2txt.process(BytesIO(file_contents))

        # Extract questions from the text
        questions = []
        for line in text.split("\n"):
            line = line.strip()
            if line.endswith("?"):
                questions.append(line)

        # Define the question index and TTS engine
        question_index =  st.number_input("Current question index:",
                                                 min_value=0, max_value=len(questions) - 1, value=0, step=1)
        engine = pyttsx3.init()

        voices = engine.getProperty('voices') 
        engine.setProperty('voice', voices[0].id)  
        engine.setProperty('rate', 150)    

        # Speak the current question and display the "Speak next question" button
        st.write("### Extracted questions:")
        st.write(questions[question_index])

        # Speak the current question and display the "Speak next question" button
        if st.button("Speak current question"):
            if question_index < len(questions):
                if engine._inLoop:  # Check if the engine is already running
                    stop_button = st.button("Stop speaking")
                    engine.say(questions[question_index])
                    engine.runAndWait()
                    if stop_button:
                        stop_engine(engine)
                        st.stop()
                else:
                    engine.say(questions[question_index])
                    engine.runAndWait()

                # Initialize the recognizer and the microphone
                recognizer = sr.Recognizer()
                mic = sr.Microphone()

                # Record the answer
                audio = audio_recorder()

                # Convert speech to text
                try:
                    answer = recognizer.recognize_google(audio)
                    st.write("### Answer:")
                    st.write(answer)
                    if engine._inLoop:  # Check if the engine is already running
                        stop_button = st.button("Stop speaking")
                        engine.say("Answer: " + answer)
                    engine.runAndWait()

                    st.session_state.answers[question_index] = answer  # Store the answer in the session_state.answers dictionary

                except sr.UnknownValueError:
                    st.write("Speech Recognition could not understand the audio.")
                except sr.RequestError:
                    st.write("An error occurred while trying to fetch results from the Speech Recognition service.")

        # Save answers to the document
        if st.button("Save answers to the document"):

            if uploaded_file is not None:  # Check if the file is uploaded
                # Load the Word document
                doc = docx.Document(BytesIO(file_contents))

                # Add the title to the document
                title = doc.add_paragraph("Question and Answer")
                title.runs[0].font.size = docx.shared.Pt(24)
                title.runs[0].font.bold = True
                doc.add_paragraph("")  # Add an empty line after the title

                # Add the answers to the document
                for index, answer in st.session_state.answers.items():
                    question_paragraph = doc.add_paragraph()
                    question_paragraph.add_run("Question {}: ".format(index + 1)).bold = True
                    question_paragraph.add_run(questions[index])
                    answer_paragraph = doc.add_paragraph()
                    answer_paragraph.add_run("Answer {}: ".format(index + 1)).bold = True
                    answer_paragraph.add_run(answer)

                # Save the document with answers
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

