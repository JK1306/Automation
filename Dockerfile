FROM ubuntu:18.04
WORKDIR /app

RUN apt-get --fix-missing update && \
    apt-get install -y software-properties-common && \
    rm -rf /var/lib/apt/lists/*

RUN apt-get update 

RUN add-apt-repository universe

RUN apt-get install firefox xvfb -y
RUN apt-add-repository ppa:mozillateam/firefox-next
RUN apt install firefox xvfb -y
RUN apt-get update --fix-missing && apt-get -y install python3-pip


RUN pip3 install virtualenv

RUN virtualenv main_api -p python3
RUN cd main_api && . ./bin/activate
RUN pip3 install gunicorn
COPY requirements.txt /app/main_api

RUN pip3 install -r /app/requirements.txt


RUN mkdir -p /app/main_api/SPI_Group

COPY ./SPI_Group/OLD /app/main_api/SPI_Group
RUN DEBIAN_FRONTEND="noninteractive" apt-get -y install tzdata
RUN apt-get install vim -y
RUN ls -l
WORKDIR /app/main_api/SPI_Group
CMD ["python3", "task_2.py"]

