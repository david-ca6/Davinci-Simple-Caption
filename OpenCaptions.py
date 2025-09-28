#!/usr/bin/env python3

import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ------------------------- resolve api connection -------------------------

try:
    resolve
except NameError:
    from python_get_resolve import GetResolve
    resolve = GetResolve()

project_manager = resolve.GetProjectManager()
project = project_manager.GetCurrentProject()
timeline = project.GetCurrentTimeline()

# ------------------------- srt file functions -------------------------

def srt2df(file_path):
    df = []

    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()

    subtitle_blocks = content.strip().split('\n\n')

    for block in subtitle_blocks:
        lines = block.split('\n')
        if len(lines) >= 3:

            nid = int(lines[0])

            timestamp = lines[1]
            text = '\n'.join(lines[2:])

            start_time = timestamp.split(' --> ')[0]
            end_time = timestamp.split(' --> ')[1]
            
            h, m, s = start_time.replace(',', '.').split(':')
            startTime_seconds = float(h) * 3600 + float(m) * 60 + float(s)

            h, m, s = end_time.replace(',', '.').split(':')
            endTime_seconds = float(h) * 3600 + float(m) * 60 + float(s)

            df.append({'id': nid, 'start': startTime_seconds, 'end': endTime_seconds, 'text': text})

    return df

def df2srt(df, file_path):
    with open(file_path, 'w', encoding='utf-8') as file:
        for row in df:

            nid = row['id']

            if nid == 0:
                continue
            
            start_time = timedelta(seconds=row['start'])
            hours, remainder = divmod(start_time.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            milliseconds = start_time.microseconds // 1000
            start_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d},{milliseconds:03d}"

            end_time = timedelta(seconds=row['end'])
            hours, remainder = divmod(end_time.seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            milliseconds = end_time.microseconds // 1000
            end_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d},{milliseconds:03d}"

            file.write(f"{nid}\n")
            file.write(f"{start_time_str} --> {end_time_str}\n")
            file.write(f"{row['text']}\n\n")

def remove_ponctuation(text):
    ponctuation = [".", ","]
    for ponctuation in ponctuation:
        text = text.replace(ponctuation, "")
    return text

def apply_text_transform(text, transform):
    if transform == "Lowercase":
        return text.lower()
    elif transform == "Uppercase":
        return text.upper()
    elif transform == "Capitalize All Words":
        return text.title()
    else:
        return text

# ------------------------- resolve timeline functions -------------------------

def timelineText2df(timeline, marker):
    df = []
    if timeline:
        track_count = timeline.GetTrackCount("video")
        for i in range(1, track_count + 1):
            track = timeline.GetItemListInTrack("video", i)
            if track:
                track_name = timeline.GetTrackName("video", i)
                if track_name == marker:
                    nid = 1
                    for item in track:
                        if item.GetName() == "Text+":
                            start_time = item.GetStart() / timeline.GetSetting('timelineFrameRate')
                            end_time = item.GetEnd() / timeline.GetSetting('timelineFrameRate')
                            fusion_comp = item.GetFusionCompByIndex(1)
                            if fusion_comp:
                                text_tool = fusion_comp.FindToolByID("TextPlus")
                                if text_tool:
                                    text_content = text_tool.GetInput("StyledText")
                                    df.append({'id': nid, 'start': start_time, 'end': end_time, 'text': text_content})
                                    nid += 1
    return df

# ------------------------- srt file functions -------------------------

def df2timelineText(df, timeline, marker):
    if timeline:
        track_count = timeline.GetTrackCount("video")
        for i in range(1, track_count + 1):
            track = timeline.GetItemListInTrack("video", i)
            if track:
                track_name = timeline.GetTrackName("video", i)
                nid = 1
                if track_name == marker:
                    for item in track:
                        if item.GetName() == "Text+":
                            start_time = item.GetStart() / timeline.GetSetting('timelineFrameRate')
                            end_time = item.GetEnd() / timeline.GetSetting('timelineFrameRate')
                            fusion_comp = item.GetFusionCompByIndex(1)
                            if fusion_comp:
                                text_tool = fusion_comp.FindToolByID("TextPlus")
                                if text_tool:
                                    for row in df:
                                        if row['id'] == nid:
                                            text_tool.SetInput("StyledText", row['text'])
                                            nid += 1
                                            break

def df2NewtimelineText(df, timeline, template_name, remove_punctuation=True, text_transform="None"):
    """
    Create new Text+ clips from SRT dataframe on timeline
    
    Args:
        df: list of dicts with subtitle data
        timeline: DaVinci Resolve timeline object
        template_name: Name of the Text+ template to use
        remove_punctuation: Whether to remove punctuation from text
    """
    if not timeline or not df:
        print("No timeline or empty dataframe")
        return False

    print(f"Creating Text+ clips from SRT file: {df} using template: {template_name}")
    
    media_pool = project.GetMediaPool()
    root_folder = media_pool.GetRootFolder()
    
    text_clip = find_text_plus_template_by_name(media_pool, template_name)    
    if not text_clip:
        print(f"Text+ template '{template_name}' not found in Media Pool.")
        print("Available templates:")
        list_available_templates(media_pool)
        return False
    
    print(f"Found Text+ template: {text_clip.GetClipProperty('Clip Name')}")
    
    track_added = timeline.AddTrack("video")
    if not track_added:
        print("Failed to add new video track")
        return False
    
    track_count = timeline.GetTrackCount("video")
    
    fps = float(timeline.GetSetting('timelineFrameRate'))
    
    duration_multiplier = 1.0
    try:
        test_duration = 100
        test_clip = {
            "mediaPoolItem": text_clip,
            "startFrame": 0,
            "endFrame": test_duration - 1,
            "trackIndex": track_count,
            "recordFrame": 0
        }
        
        test_items = media_pool.AppendToTimeline([test_clip])
        if test_items and len(test_items) > 0:
            test_item = test_items[0]
            test_duration_real = test_item.GetDuration()
            timeline.DeleteClips([test_item], False)
            duration_multiplier = test_duration / test_duration_real if test_duration_real > 0 else 1.0
        
        print(f"Duration multiplier: {duration_multiplier:.3f}")
    except Exception as e:
        print(f"Warning: Could not calculate duration multiplier, using 1.0: {e}")
        duration_multiplier = 1.0
    
    created_clips = []
    
    for row in df:
        if row['id'] == 0:
            continue
            
        start_frame = int(row['start'] * fps)
        end_frame = int(row['end'] * fps)
        duration = end_frame - start_frame
        
        new_clip = {
            "mediaPoolItem": text_clip,
            "startFrame": 0,
            "endFrame": duration - 1,
            "trackIndex": track_count,
            "recordFrame": start_frame
        }
        
        base_duration = new_clip["endFrame"] - new_clip["startFrame"] + 1
        new_duration = int(base_duration * duration_multiplier + 0.999)
        new_clip["endFrame"] = new_duration - 1
        timeline_items = media_pool.AppendToTimeline([new_clip])
        
        if timeline_items and len(timeline_items) > 0:
            timeline_item = timeline_items[0]
            
            timeline_item.SetClipColor("Green")

            if timeline_item.GetFusionCompCount() > 0:
                comp = timeline_item.GetFusionCompByIndex(1)
                if comp:
                    text_tool = comp.FindToolByID("TextPlus")
                    if text_tool:
                        text_content = remove_ponctuation(row['text']) if remove_punctuation else row['text']
                        text_content = apply_text_transform(text_content, text_transform)
                        text_tool.SetInput("StyledText", text_content)
                        created_clips.append(timeline_item)
                        print(f"Created subtitle {row['id']}: {row['text'][:50]}{'...' if len(row['text']) > 50 else ''}")
                    else:
                        print(f"Warning: No TextPlus tool found in template for subtitle {row['id']}")
            else:
                print(f"Warning: No Fusion composition found for subtitle {row['id']}")
        else:
            print(f"Error: Failed to create timeline item for subtitle {row['id']}")

    
    print(f"Created {len(created_clips)} Text+ clips")
    return True

def find_text_plus_template_by_name(media_pool, template_name):
    """
    Find a specific Text+ template by name in the media pool
    Searches all folders for a matching template name
    """
    root_folder = media_pool.GetRootFolder()
    
    def search_folder(folder):
        clips = folder.GetClipList()
        for clip in clips:
            if clip.GetClipProperty("File Path") == "":
                clip_name = clip.GetClipProperty("Clip Name")
                if clip_name == template_name:
                    return clip
        
        for subfolder in folder.GetSubFolderList():
            result = search_folder(subfolder)
            if result:
                return result
        return None
    
    return search_folder(root_folder)

def list_available_templates(media_pool):
    """
    List all available Text+ templates (Fusion compositions) in the media pool
    """
    root_folder = media_pool.GetRootFolder()
    templates = []
    
    def search_folder(folder, folder_path=""):
        clips = folder.GetClipList()
        for clip in clips:
            if clip.GetClipProperty("File Path") == "":
                clip_name = clip.GetClipProperty("Clip Name")
                templates.append(f"  - {clip_name} (in {folder_path or 'Root'})")
        
        for subfolder in folder.GetSubFolderList():
            subfolder_name = subfolder.GetName()
            new_path = f"{folder_path}/{subfolder_name}" if folder_path else subfolder_name
            search_folder(subfolder, new_path)
    
    search_folder(root_folder)
    
    if templates:
        for template in templates:
            print(template)
    else:
        print("  No Text+ templates (Fusion compositions) found in Media Pool")
        print("  Create a Text+ composition in Fusion and save it to the Media Pool")

    # ------------------------------------------------------------

def get_video_tracks():
    global timeline
    timeline = project.GetCurrentTimeline()
    track_count = timeline.GetTrackCount("video")
    return [timeline.GetTrackName("video", i) for i in range(1, track_count + 1)]

def get_available_templates():
    """Get list of available Text+ templates from Media Pool"""
    try:
        media_pool = project.GetMediaPool()
        root_folder = media_pool.GetRootFolder()
        templates = []
        
        def search_folder(folder):
            for subfolder in folder.GetSubFolderList():
                if subfolder.GetName() == "Captions Templates":
                    clips = subfolder.GetClipList()
                    for clip in clips:
                        if clip.GetClipProperty("File Path") == "":
                            clip_name = clip.GetClipProperty("Clip Name")
                            templates.append(clip_name)
        
        search_folder(root_folder)

        templates.sort()

        return templates
    except Exception as e:
        print(f"Error getting templates: {e}")
        return []

def main():
    root = tk.Tk()
    root.focus_force()
    root.title("OpenCaptions")
    root.geometry("720x420")

    status_var = tk.StringVar()
    remove_punctuation_var = tk.BooleanVar(value=True)
    text_transform_options = ["None", "Lowercase", "Uppercase", "Capitalize All Words"]
    text_transform_var = tk.StringVar(value=text_transform_options[0])

    style = ttk.Style(root)
    style.configure("Delete.TButton", foreground="red")

    templates = get_available_templates()
    track_entries = []
    add_button = None

    def select_srt_file(entry):
        if entry not in track_entries:
            return
        path = filedialog.askopenfilename(title="Select SRT File", filetypes=[("SRT files", "*.srt"), ("All files", "*.*")])
        if path:
            entry["srt_var"].set(path)

    def remove_track_entry(entry):
        nonlocal add_button
        if entry not in track_entries:
            return
        for widget in (entry["label"], entry["template_combo"], entry["srt_entry"], entry["select_button"], entry["delete_button"]):
            widget.destroy()
        track_entries.remove(entry)
        for index, entry_item in enumerate(track_entries):
            entry_item["label"].configure(text=f"Track {index + 1}")
            entry_item["label"].grid_configure(row=index + 1)
            entry_item["template_combo"].grid_configure(row=index + 1)
            entry_item["srt_entry"].grid_configure(row=index + 1)
            entry_item["select_button"].grid_configure(row=index + 1)
            entry_item["delete_button"].grid_configure(row=index + 1)
        if add_button is not None:
            add_button.state(["!disabled"])
        status_var.set("Removed track.")

    def add_track_entry():
        nonlocal add_button
        if len(track_entries) >= 5:
            status_var.set("Maximum of 5 tracks.")
            if add_button is not None:
                add_button.state(["disabled"])
            return
        index = len(track_entries)
        template_var = tk.StringVar(value=templates[0] if templates else "")
        srt_var = tk.StringVar()
        entry = {
            "template_var": template_var,
            "srt_var": srt_var,
        }
        label = ttk.Label(tracks_frame, text=f"Track {index + 1}")
        label.grid(row=index + 1, column=0, sticky="w", pady=(4, 0))
        entry["label"] = label
        template_combo = ttk.Combobox(tracks_frame, textvariable=template_var, values=templates, state="readonly")
        template_combo.grid(row=index + 1, column=1, sticky="ew", pady=(4, 0))
        entry["template_combo"] = template_combo
        srt_entry = ttk.Entry(tracks_frame, textvariable=srt_var)
        srt_entry.grid(row=index + 1, column=2, sticky="ew", pady=(4, 0))
        entry["srt_entry"] = srt_entry
        select_button = ttk.Button(tracks_frame, text="Select", command=lambda e=entry: select_srt_file(e))
        select_button.grid(row=index + 1, column=3, sticky="w", pady=(4, 0))
        entry["select_button"] = select_button
        delete_button = ttk.Button(tracks_frame, text="X", width=1, style="Delete.TButton", command=lambda e=entry: remove_track_entry(e))
        delete_button.grid(row=index + 1, column=4, sticky="w", pady=(4, 0))
        entry["delete_button"] = delete_button
        track_entries.append(entry)
        if len(track_entries) == 5 and add_button is not None:
            add_button.state(["disabled"])

    def refresh_templates():
        nonlocal templates
        templates = get_available_templates()
        if templates:
            status_var.set(f"Found {len(templates)} templates")
        else:
            status_var.set("No Text+ templates found in Media Pool")
        for entry in track_entries:
            entry["template_combo"]["values"] = templates
            if templates:
                if entry["template_var"].get() not in templates:
                    entry["template_var"].set(templates[0])
            else:
                entry["template_var"].set("")

    def execute_callback():
        if not track_entries:
            status_var.set("Add at least one track.")
            return
        for index, entry in enumerate(track_entries, start=1):
            if not entry["srt_var"].get() or not entry["template_var"].get():
                status_var.set(f"Track {index} is missing an SRT file or template.")
                return
        global timeline
        timeline = project.GetCurrentTimeline()
        for index, entry in enumerate(track_entries, start=1):
            df = srt2df(entry["srt_var"].get())
            success = df2NewtimelineText(
                df,
                timeline,
                entry["template_var"].get(),
                remove_punctuation=remove_punctuation_var.get(),
                text_transform=text_transform_var.get(),
            )
            if not success:
                status_var.set(f"Failed to create track {index}.")
                return
        status_var.set(f"Created {len(track_entries)} Text+ tracks.")

    content = ttk.Frame(root, padding=16)
    content.grid(row=0, column=0, sticky="nsew")
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)

    content.columnconfigure(0, weight=1)
    content.columnconfigure(1, weight=1)
    content.columnconfigure(2, weight=1)
    content.rowconfigure(1, weight=1)

    ttk.Label(content, text="Tracks").grid(row=0, column=0, columnspan=3, sticky="w")

    tracks_frame = ttk.Frame(content)
    tracks_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
    tracks_frame.columnconfigure(1, weight=1)
    tracks_frame.columnconfigure(2, weight=1)

    ttk.Label(tracks_frame, text="Track").grid(row=0, column=0, sticky="w")
    ttk.Label(tracks_frame, text="Template").grid(row=0, column=1, sticky="w")
    ttk.Label(tracks_frame, text="SRT File").grid(row=0, column=2, sticky="w")
    ttk.Label(tracks_frame, text="Load File").grid(row=0, column=3, sticky="w")
    ttk.Label(tracks_frame, text="Remove").grid(row=0, column=4, sticky="w")

    controls_frame = ttk.Frame(content)
    controls_frame.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
    add_button = ttk.Button(controls_frame, text="Add Track", command=add_track_entry)
    add_button.grid(row=0, column=0, sticky="w")
    ttk.Button(controls_frame, text="Refresh Templates", command=refresh_templates).grid(row=0, column=1, sticky="w")
    controls_frame.columnconfigure(2, weight=1)

    add_track_entry()
    refresh_templates()

    ttk.Label(content, text="Case conversion").grid(row=3, column=0, sticky="w", pady=(16, 0))
    transform_combo = ttk.Combobox(content, textvariable=text_transform_var, values=text_transform_options, state="readonly")
    transform_combo.grid(row=3, column=1, sticky="ew", pady=(16, 0))
    ttk.Checkbutton(content, text="Remove punctuation", variable=remove_punctuation_var, onvalue=True, offvalue=False).grid(row=4, column=0, columnspan=3, sticky="w", pady=(12, 0))
    ttk.Button(content, text="Execute", command=execute_callback).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(16, 0))
    status_lbl = ttk.Label(content, textvariable=status_var)
    status_lbl.grid(row=6, column=0, columnspan=3, sticky="w", pady=(12, 0))

    root.mainloop()

if __name__ == "__main__":
    main()