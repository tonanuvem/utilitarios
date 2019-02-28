FROM python:3.6-alpine
ADD . /code
WORKDIR /code

RUN apt-get install python-dev libxml2-dev libxslt1-dev antiword unrtf poppler-utils pstotext tesseract-ocr \
flac ffmpeg lame libmad0 libsox-fmt-mp3 sox libjpeg-dev swig

RUN pip3 install textract
# RUN pip3 install -r requirements.txt

CMD ["python", "word/word-to-txt.py"]
