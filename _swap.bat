@ECHO OFF
call "C:\Users\Dealing\PycharmProjects\MIDSHIFT\venv\Scripts\activate.bat"
cd C:\Users\Dealing\PycharmProjects\MIDSHIFT
streamlit run test.py


start "" http://localhost:8501/
pause
