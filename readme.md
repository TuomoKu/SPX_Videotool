﻿﻿﻿﻿﻿﻿# SPX Videotool

is a scripted Windows utility for automated videofile conversions.
(Available at https://bitbucket.org/TuomoKu/spx-videotool)

The basic principle is that you create **Sourcefolders**, assign **Conversion Tasks** to them and the resulted new videofile(s) will be moved to **Target Folders**. This process repeats periodically.

All of these are configured in **config.ini** -file.

![SPX Videotool UI](https://bitbucket.org/TuomoKu/spx-videotool/downloads/SPX_VideotoolUI.png "SPX Videotool UI")



---


## Important concepts:
- **converter**: an executable program which will do the processing part, such as *ffmpeg*.
- **interval**: the delay time in seconds between folder scan processes.
- **config.ini**: a file where all tasks and folders are configured.
- **task** is the indivual conversion job and it's folders and other parameters. You can have unlimited, named tasks in the config file.
- **source / target folder** Each task has a source folder which is observed for new "incoming video files" and the processed outcome will be saved to the tasks target folder.
- **special variables** are built in variables for helping configuring tasks. Variables uses _double mustache_ format. The following variables are available
>- **{{SOURCEFILE}}** a runtime evaluated full path to the source file (for example "c:/source/videofile.avi")
>-  **{{TARGETFILE}}** a runtime evaluated path to the target file without extension ("c:/source/videofile")
>- **{{EXTENSION}}** file extension of the source file ("avi")
>- **{{ROOTFOLDER}}** refers to the folder containing the videotool application (SPX_VIDEOTOOL.hta)

---

## Installation and usage
- Download required files (also available as a zip-file: https://drive.google.com/file/d/1PswRjMxGYkhnU-_qBHS4CwiqiAMPaFhr/view?usp=sharing ) and extract to a new folder, such as "C:/Videotool".
- Go to the folder and create a suitable sub-folder structure for tasks, see example below.
- edit config.ini to match your folder structure and other needs.
- Double click on SPX_VIDEOTOOL.hta and the UI opens up and starts working.
- Logs are generated to _log_ -subfolder


## Example folder structure
```
VIDEOTOOL
* SPX_VIDEOTOOL.hta
* config.ini
>- src
>- log
v- bin
    * ffmpeg.exe
v- TASKS
    v- EXTRACT_AUDIO
        >- OUT
        >- PROCESSED_FILES
        >- FAILED_FILES
    v- MAKE_PREVIEW
        * source_waiting_1.mov
        * source_waiting_2.mov
        v- OUT
              * targetfile.mov
        v- PROCESSED_FILES
              * sourcefile.mov
        v- FAILED_FILES
              * problematic_sourcefile.mov




```

## Example task in config.ini
```
[H264 PREVIEW]
Task_Description = Make 720p h264 mp4 preview file.
Source_Directory = {{ROOTFOLDER}}\VIDEO_TASKS\MAKE_PREVIEW\
Source_FileExtns = mov, avi, mxf, mp4
Target_Directory = {{ROOTFOLDER}}\VIDEO_TASKS\MAKE_PREVIEW\TARGET\
Converter_Progrm = {{ROOTFOLDER}}\bin\ffmpeg.exe
Converter_Params = -y -i {{SOURCEFILE}} -vcodec h264 -acodec mp2 -s 1280x720 {{TARGETFILE}}_h264_720p.mp4
```


## "Software" "architecture" :-)

### GUI
The main file is **SPX_VIDEOTOOL.hta** which basically is special HTML-file which will load a customized web browser which comes with Windows. This browser can load local VB Scripts which are not limited to browser but can work on operating system level just as any other VB Script can. Therefor it can process file system level tasks, access databases, read and write local files, start applications, access internet, etc. Basically HTA can be considered a graphical user interface (GUI) on top of VB Script.

Read more about HTA at https://en.wikipedia.org/wiki/HTML_Application

### LOGIC
Majority of the program logic is written in Visual Basic Script (vbs) language. VB Scripts can access system resources, read and write files and execute programs.  VB Scripts have somewhat notorious reputation as being source of malicious apps of many kinds, but SPX Videotool here is an example of the useful stuff you can do with it :-) Read more about VBS at https://en.wikipedia.org/wiki/VBScript


### VIDEO CONVERTER
The package comes with **ffmpeg** binary, which is a command line utility for video file processing (see http://https://www.ffmpeg.org/). Any other command line procesor can be used also, suck as Imagemagick. SPX Videotool creates command line arguments for the converter and starts the process. 
The converter application can be configured separately for each task with _config.ini._

---

## MIT License
Copyright 2020 Tuomo Kulomaa <tuomo@smartpx.fi>

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. 

THE SOFTWARE IS PROVIDED "AS IS", **WITHOUT WARRANTY OF ANY KIND**, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 

---

























