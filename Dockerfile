FROM python:3.9

# set environment variables
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

COPY requirements.txt .

# install python dependencies
RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

RUN wget -nc https://dl.winehq.org/wine-builds/winehq.key
RUN apt-key add winehq.key
RUN add-apt-repository 'deb https://dl.winehq.org/wine-builds/ubuntu/ focal main'
RUN apt-get update
RUN apt-get install -y winehq-stable

RUN apt-get install -y winetricks

RUN apt-get clean -y
RUN apt-get autoremove -y

ENV WINEDEBUG=fixme-all

RUN winetricks msxml6


# gunicorn
CMD ["gunicorn", "--config", "gunicorn-cfg.py", "run:app"]
