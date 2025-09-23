# Davinci Simple Caption

This is a simple tool that uses the Davinci API to create text+ caption track on a Davinci Timeline using a .srt files and a Text+ template.

- Simple Caption always work with the focused timeline, no need to restart it when you change the timeline.
- Simple Caption will always create a new text+ track, it will not overwrite existing text+ tracks.

## Setup
1. Install the dependencies: 
```
pip install pandas dearpygui
```

2. Create a "Captions Templates" folder in your Media Pool. And place your Text+ templates in it.

## Usage

1. Create write or generate your subtitle.

2. Export your subtitle to a .srt file.

3. Open Simple Caption.
```
python SimpleCaption.py
```

4. Select your SRT file and your Text+ template.

5. Click "Execute".

6. Enjoy!

## Dependencies

- Python 3.10+
- pandas
- dearpygui

## About
This project is based on my older project [Resolve_TextPlus2SRT](https://github.com/david-ca6/Resolve_TextPlus2SRT), but simpler and more user-friendly.

