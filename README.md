# Embroidery Combiner

Combines multiple embroidery design files (.ngs, .dst) into a single file, stacked vertically with a configurable gap between each design.

## How to Use

1. **Open the app** — Double-click `EmbroideryC.exe`
2. **Select a folder** — Click **Browse** and pick the folder containing your embroidery files
3. **Check your files** — The file list shows every design found. Uncheck any you want to skip
4. **Set the gap** — Choose a preset (Tight, Normal, Wide, Extra Wide) or type a custom value in mm
5. **Combine** — Click **Combine Files**. The combined `.dst` file is saved in the same folder

## What Each Section Does

| Section | What it does |
|---------|-------------|
| **Browse** | Pick the folder with your .ngs or .dst files |
| **File list** | Shows all embroidery files found. Check/uncheck to include or skip files |
| **Gap between designs** | Space (in mm) between each design when stacked vertically |
| **Save as** | The output filename. Auto-generated from file numbers (e.g. `216-225.dst`) |
| **Combine Files** | Runs the combine. Progress bar shows each step |
| **Settings** | Set the path to Wings XP or My Editor (for NGS conversion), change theme |

## NGS Files (Windows Only)

`.ngs` files need to be converted to `.dst` before combining. The app does this automatically using **Wings XP** or **My Editor** — but only on Windows.

If the app can't find Wings/My Editor automatically:
1. Click **Settings**
2. Click **Browse** next to the editor path
3. Navigate to the Wings or My Editor `.exe` file
4. Click **Save**

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "No embroidery files found" | Make sure the folder contains `.ngs` or `.dst` files |
| "Wings Editor Not Found" | Open Settings and set the path to Wings/My Editor manually |
| "Windows Required" | NGS files can only be converted on Windows. Convert to DST first, or run the app on a Windows machine |
| "Permission Denied" | Close the file in any other program and try again |
| File exists warning | The output file already exists. You can replace it or change the filename |

## Version

Check the current version in **Settings** (bottom-left corner of the settings dialog).

Latest releases: [GitHub Releases](https://github.com/shaanpawa/embroidery-combiner/releases)
