# Davinci Simple Caption
A Simple and Open Source Caption Tool for Davinci Resolve.

## Description
This is an open source tool that uses the Davinci API to create text+ caption track on a Davinci Timeline using a .srt files and a Text+ template.

- Simple Caption always work with the focused timeline, no need to restart it when you change the timeline.
- Simple Caption will always create a new text+ track, it will not overwrite existing text+ tracks.

## Setup
1. Install [DaVinci Resolve Studio](https://www.blackmagicdesign.com/products/davinciresolve) 19 or higher.
2. Install [Python](https://www.python.org/downloads/) 3.10 or higher.
3. Create a "Captions Templates" folder in your Media Pool. And place your Text+ templates in it.

## Usage

1. Create write or generate your subtitle.

2. Export your subtitle to a .srt file.

3. Open Simple Caption.
```
chmod +x ./run.sh
./run.sh
```

4. Select your SRT file and your Text+ template.

5. Click "Execute".

6. Enjoy!

## Dependencies

- Python 3.10+
- tkinter (standard library)

## About
This project is based on the code from my older project [Resolve_TextPlus2SRT](https://github.com/david-ca6/Resolve_TextPlus2SRT), but simpler to use and more user-friendly, because not everyone want to run a script in the terminal. 
