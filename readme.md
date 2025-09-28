# Davinci Simple Caption
A Simple and Open Source Caption Tool for Davinci Resolve.

## Description
This is an open source tool that uses the Davinci API to create text+ caption track on a Davinci Timeline using a .srt files and a Text+ template.

- Simple Caption always work with the focused timeline, no need to restart it when you change the timeline.
- Simple Caption will always create a new text+ track, it will not overwrite existing text+ tracks.

## Features
- Create Text+ from a .srt file and a Text+ template.
- Remove punctuation (optional)
- Case conversion [none lover case, upper case, capitalize all words]

## Setup
1. Install [DaVinci Resolve](https://www.blackmagicdesign.com/products/davinciresolve) 19 or higher.
2. Install [Python](https://www.python.org/downloads/) 3.10 or higher.

## Usage as Resolve Script (Davinci Resolve Studio and Davinci Resolve Free)
1. Install Simple Caption 
2. Create a "Captions Templates" folder in your Media Pool. 
3. Place your Text+ templates in it.
4. Create write or generate your subtitle.
5. Export your subtitle to a .srt file.
6. Run Simple Caption from the resolve Workspace menu. ` Workspace -> Scripts -> Comp -> SimpleCaption`
7. Select your SRT file and your Text+ template.
8. Click "Execute".
9. Enjoy!

## Usage as External Script (Davinci Resolve Studio Only)
1. Create a "Captions Templates" folder in your Media Pool. 
2. Place your Text+ templates in it.
2. Create write or generate your subtitle.
3. Export your subtitle to a .srt file.
4. Run Simple Caption.py from a terminal.
5. Select your SRT file and your Text+ template.
6. Click "Execute".
7. Enjoy!

## Why Simple Use Simple Caption?
- Simple to use
- Totaly free
- Totally open source, you can audit the code, and make your own changes
- Cross-platform, you can use it on Windows, macOS and Linux
- Compatible with both Davinci Resolve Free and Davinci Resolve Studio (paid)

## Dependencies
- Python 3.10+
- tkinter (standard library)

## About
The starting point of Simple Caption is based on one of my older project [Resolve_TextPlus2SRT](https://github.com/david-ca6/Resolve_TextPlus2SRT). 
But TextPlus2SRT was more a custom script for my own use than anything else, it was missing a lot of features, it was only working with linux, required to type in a terminal, and it only allowed to convert SRT to TextPlus, nothing more. Simple Caption is inteded to be a stronger base to work from to make a more powerful and user-friendly tool.

