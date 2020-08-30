FROM ubuntu:18.04

RUN apt-get --fix-missing update && \
    apt-get install -y software-properties-common && \
    rm -rf /var/lib/apt/lists/*

RUN apt-get update 

RUN add-apt-repository universe
RUN apt-get update

RUN apt-get install firefox xvfb -y
RUN apt-add-repository ppa:mozillateam/firefox-next
RUN apt install firefox xvfb -y
RUN apt-get update --fix-missing && apt-get -y install python3-pip



WORKDIR /app
RUN virtualenv main_api -p python3
RUN cd main_api && . ./bin/activate
RUN pip3 install gunicorn
COPY requirements.txt /app/main_api

RUN pip3 install -r /app/requirements.txt


RUN mkdir -p /app/main_api/SPI_Group

ENV PYTHONIOENCODING UTF-8
RUN DEBIAN_FRONTEND="noninteractive" apt-get -y install tzdata
RUN apt-get install python3-tk -y
RUN apt-get install vim -y
COPY ./ /app/
RUN ls -l
RUN cd rap-bot && python3 setup.py build
RUN cd rap-bot && python3 setup.py install
ENV PATH=$PATH:/app/.

RUN pwd
CMD ls -l
CMD ["python3", "-u", "final_firefox_latest.py"]

