# Solar System Explorer

*An Interactive Excel Workbook for Exploring the Solar System*  
**Project Version 1.2.0 — September 25, 2025**
**Source Version 1.1.0 — September 25, 2025**

---

## Overview

*Solar System Explorer* is a macro-enabled Excel workbook designed to organize, sort, and visualize data about objects in the Solar System. It combines standardized datasets with custom VBA macros to provide interactive tools for exploring planetary and minor body characteristics. This project is intended as both a reference and an interactive exploration tool, offering users the ability to sort, recolor, and visualize the Solar System dynamically within Excel.

Inclusion standards for the workbook are as follows:

- All planets and dwarf planets  
- All known natural satellites  
- Major NEOs, plus those visited by spacecraft  
- Asteroids, Centaurs, and Jupiter Trojans ≥ 50 km radius  
- All known Trojan asteroids of the other planets  
- KBOs, SDOs, and detached objects ≥ 100 km radius  
- Select asteroids and TNOs visited by spacecraft  

---

## Contents

I. [Installation / Setup](#i-installation--setup)  
II. [Macros & Features](#ii-macros--features)  
III. [Using the Workbook](#iii-using-the-workbook)  
IV. [Known Limitations / Notes](#iv-known-limitations--notes)  
V. [Credits & Version History](#v-credits--version-history)  

---

## I. Installation / Setup

1. **Download the files**  
   - Go to the [Releases page](https://github.com/amorphose/Solar-System-Explorer/releases) and download the latest version of *Solar System Explorer*.  
   - You can download the workbook directly (`.xlsm`) or use the “Source code (zip)” which also includes this README and LICENSE.  
   - The required font, [Astromoony](https://github.com/RobertWinslow/Astromoony-Font), must be downloaded separately.  

2. **Unblock the file to allow macros**  
   - Right-click on *Solar System Explorer`.xlsm`* and choose **Properties**.  
   - If you see a message at the bottom that says *“This file came from another computer and might be blocked to help protect this computer”*, check **Unblock**, then click **Apply**.  
   - This ensures that Excel will not automatically disable macros.  

3. **Install the Astromoony font**  
   - Download the [Astromoony font](https://github.com/RobertWinslow/Astromoony-Font) if you haven’t already.  
   - Right-click the `.ttf` file and choose **Install** (or **Install for all users**).  
   - Once installed, the workbook will display special dwarf planet symbols correctly.  

4. **Enable macros in Excel**  
   - When opening the workbook for the first time, you may see a security warning about macros.  
   - Click **Enable Content** to allow the macros to run. The workbook’s sorting features will not function correctly without macros enabled.  

5. **Add the “Sorting Tools” custom tab** (optional but recommended)  
   - Go to **File > Options > Customize Ribbon**.  
   - On the right panel, click **New Tab**, then rename it “Sorting Tools.”  
   - With the new tab selected, click **New Group** to create a group inside it.  
   - In the left panel, change the dropdown from “Popular Commands” to **Macros**.  
   - Select the following macros and add them to your new group:  
     - `MirrorSortingData`  
     - `ResetSortingData`  
     - `ColorizeSolarSystem`  
   - (Tip: You can also click **Rename...** to add spaces and assign custom icons for easier recognition.)  
   - Finally, use the up/down arrows to position the **Sorting Tools** tab between **Cells** and **Editing** on the ribbon. This places the macros directly next to the **Custom Sort** command.  
   - Click **OK** to save. The “Sorting Tools” tab will now appear on your Excel ribbon.  

---

## II. Macros & Features

The workbook relies on VBA macros to extend Excel’s native functionality. There are two types: **user-facing macros** you can run directly, and **background macros** that run automatically.

**User-Facing Macros** (via the “Sorting Tools” ribbon tab):  
- **Mirror Sorting Data** – Updates the Solar System sheet so its order matches the Sorting Data sheet.  
- **Reset Sorting Data** – Restores the Solar System and Sorting Data sheets to their original order.  
- **Colorize Solar System** – Reapplies domain color coding after a mirror sort.  

**Background Macros** (run automatically in the Orbital Plotter):  
- **Resize Orbital Plot** & **Resize Satellite Plot** – Adjusts chart scales based on aphelion distances.  
- **Colorize Orbital Plot** & **Colorize Satellite Plot** – Colors orbital paths by domains.  
- **Format Special Symbol** – Ensures dwarf planet symbols display correctly.  

Together, these macros provide consistent colorization, accurate orbital plots, and flexible sorting of Solar System data.  

---

## III. Using the Workbook

The workbook contains four main sheets:

1. **Solar System**  
   - Objects start sorted by distance from the Sun.  
   - Use **Mirror Sorting Data** to apply custom order and formatting from **Sorting Data**.  
   - Use **Colorize Solar System** to restore alternating row colors for readability.  

2. **Sorting Data**  
   - Start here to change the sort order of objects.  
   - Use Excel’s **Custom Sort**, then return to **Solar System** to mirror and re-color.  
   - Use **Reset Sorting Data** to restore default order.  

3. **Color Key**  
   - Reference guide for domain colors (planets, asteroid belt, Kuiper belt, etc.).  

4. **Orbital Plotter**  
   - Select an object to visualize its orbit.  
   - Orbits are scaled and colored automatically.  
   - Currently supports planets, dwarf planets, select asteroids and natural satellites.

---

## IV. Known Limitations / Notes

1. **Excel compatibility**  
   - Recommended for Excel 2016 or later (including Microsoft 365).  
   - Functionality is not guaranteed in Excel Online or Google Sheets.  

2. **Macros required for sorting**  
   - Custom sorting features will not function without macros enabled.  
   - Default views still work without macros.  

3. **Sorting caution**  
   - Always use this  workflow for custom sorts: (*Sorting Data*) **Custom Sort** → (*Solar System*) **Mirror Sorting Data** → **Colorize Solar System** (optional) → **Reset Sorting Data**.  
   - In some versions of Excel, **“My data has headers”** may need to be checked manually.  
   - Sorting **Solar System** directly will break sheet alignment.  

4. **Font requirement**  
   - Astromoony font must be installed for dwarf planet symbols to display correctly.  

5. **Custom ribbons**  
   - If the workbook is renamed or a new version is downloaded, custom ribbon buttons must be reassigned (since button references are tied to file name).  

---

## V. Credits & Version History

**Credits**  
- Workbook design & development: S. Bianchini (*amorphose*) [sbianchini@protonmail.com](mailto:sbianchini@protonmail.com)  
- Font: [Robert Winslow](https://github.com/RobertWinslow/) — Astromoony, sourced from GitHub  
- Data sources:  
  - [NASA JPL Small-Body Database](https://ssd.jpl.nasa.gov/tools/sbdb_lookup.html#/)  
  - [Mike Brown’s list of potential dwarf planets](https://web.gps.caltech.edu/~mbrown/dps.html)  
  - [Johnston’s Archive](https://www.johnstonsarchive.net/astro/index.html)  
  - [Wikipedia](https://www.wikipedia.org/)  
- Assisted with VBA coding/troubleshooting by ChatGPT-5 (OpenAI).  

#### Changelog (Version History)
**[Changelog](CHANGELOG.md)**



---

Future updates may add additional objects, expand functionality, fix errors, or refine the interface based on user feedback. Please feel free to send comments, suggestions, or bug reports to [sbianchini@protonmail.com](mailto:sbianchini@protonmail.com).
