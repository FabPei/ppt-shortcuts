# Instrumenta Keys
Instrumenta Keys is a **keyboard shortcut companion** for [Instrumenta](https://github.com/iappyx/Instrumenta/), bringing **customizable keyboard shortcuts** to Instrumenta. Since **VBA for PowerPoint lacks native keyboard shortcut support**, Instrumenta Keys -written in **Microsoft PowerShell**- runs independently alongside Instrumenta, enabling users to **assign shortcuts via a simple CSV file**.

[@iappyx](https://github.com/iappyx)

## Features
- **Fully configurable:** Easily customize shortcuts in the built-in editor to match your workflow (or import shortcut-files created by others!)
- **Runs independently:** Instrumenta Keys works alongside Instrumenta without modifying its core functionality.
- **Compatible with other VBA projects:** Designed to integrate with any VBA-based automation, not just Instrumenta.

## Experimental Notice
Instrumenta Keys is **highly experimental** and is licensed under the **MIT license**. You may freely use, modify, and distribute it, but **use at your own risk**. If you integrate this code into your own project—whether for **personal or commercial use**—please provide proper attribution in line with the **MIT license requirements**.

As stated in the license: THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

![image](https://github.com/user-attachments/assets/2962b007-77b6-4142-9c36-f9ae8886bae1)

## Platform Support
Instrumenta Keys supports **Windows** and **Mac** versions of PowerPoint, with distinct implementations for each platform:

- On **Windows**, Instrumenta Keys uses a single **PowerShell script** to connect to PowerPoint and respond to keyboard shortcuts.
- On **Mac**, Instrumenta Keys requires a helper add-in for PowerPoint that enables **AppleScript** to trigger its features via an external application.

Due to these differences, the following sections provide platform-specific details on installation, shortcut handling, and building from source.

# Windows
On Windows Instrumenta Keys is a **PowerShell script** and does not require administrative rights for installation on most enterprise systems.

## Installation
1. **Download the binary:** [Instrumenta Keys.exe](https://github.com/iappyx/Instrumenta-Keys/raw/main/bin/windows/Instrumenta%20Keys.exe)
2. Run the binary. It will automatically **generate a shortcut file for you** and **minimize to the system tray** after a few seconds. Enjoy your shortcuts!
3. To open it again, **right-click the Instrumenta Keys icon** in the system tray and click on **Show/Hide window**.

Note: A security notice may appear when running the binary, please refer to [these](https://support.microsoft.com/en-gb/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216) instructions from Microsoft to unblock Instrumenta Keys: (1) Open Windows File Explorer and go to the folder where you saved the file; (2) Right-click the file and choose Properties from the context menu; (3) At the bottom of the General tab, select the Unblock checkbox and select OK.

## How to load shortcuts
Shortcuts in Instrumenta Keys are stored in the `shortcuts.csv` file. If this file is missing when you launch Instrumenta Keys, a default version will be created automatically. You can manage and customize your shortcuts using the Shortcut Editor, which you can open in PowerPoint by pressing CTRL+SHIFT+ALT+Q. This editor allows you to modify your shortcuts, import or export shortcut files, and even load shortcut configurations directly from this GitHub page. This makes it easy to use shortcut files shared by others, such as sets tailored to specific consultancy firm standards.
Instructions and a full list of available macros in Instrumenta can be found [here](https://github.com/iappyx/Instrumenta-Keys/blob/main/instrumenta-keys-windows.md)

![image](https://github.com/user-attachments/assets/449c14ab-799e-4377-a249-f318118baddb)

## How to Build from Source
Building your own Instrumenta Keys is very simple:

### Requirements
- **Microsoft PowerShell**
- **PS2EXE module** (PowerShell-to-EXE converter)

### Steps
1. Locate the source code in `\src\windows\`
2. Run the build script in an elevated PowerShell window: `.\build.ps1`
3. The executable will be generated in `\bin\windows\Instrumenta Keys.exe`

# Mac

## Installation
To install Instrumenta Keys on Mac, you'll need three components: the helper add-in, an AppleScript file, and external software to trigger the AppleScript using keyboard shortcuts. The helper add-in integrates directly with PowerPoint, enabling Instrumenta's features within the application. External software is required to execute the AppleScript when a key combination is pressed. Due to the complex set-up the execution of shortcuts is a bit slower on Mac (takes about a second), but still functional.

### Installation of the add-in and the AppleScript-file
To install the add-in, follow these steps:

1. **Download the required files**:
   - [InstrumentaKeysHelper.ppam](https://github.com/iappyx/Instrumenta-Keys/raw/main/bin/mac/InstrumentaKeysHelper.ppam) (PowerPoint add-in)
   - [InstrumentaKeys.applescript](https://github.com/iappyx/Instrumenta-Keys/raw/main/bin/mac/InstrumentaKeys.applescript) (AppleScript helper)

2. **Move both files** to: **~/Library/Application Scripts/com.microsoft.Powerpoint/**
- This folder is located in the **Library directory of the current user**. If it does not exist, create it manually.

3. **Install the add-in**:
- Open **PowerPoint**.
- Click **Tools** in the application menu, then select **Add-ins...**.
- In the **Add-Ins** dialog box, click the **+** button, browse for the add-in file, and click **Open**.
- Click **OK** to close the dialog box.

4. **Verify the installation**:
- A **"[Keys]"** page should appear in the PowerPoint ribbon.

5. **Restart PowerPoint**:
- Fully close PowerPoint to apply the changes.
- A **reboot** may be required in some cases.
- PowerPoint may prompt you to trust the add-in or enable macros, select **"Enable Content"** or **"Enable Macros"** if prompted.

### Installation of the external shortcut handler
Several commercial applications can potentially execute the AppleScript, including Automator, BetterTouchTool, and Keyboard Maestro. While these have not been tested, they might work for triggering Instrumenta's features. However, for this installation, we are using [Karabiner-Elements](https://karabiner-elements.pqrs.org/), an open-source project that offers advanced keyboard customization. Karabiner-Elements enables users to remap keys and configure complex interactions, making it a reliable (and free!) solution for linking keyboard shortcuts to AppleScript execution.

1. **Download** Karabiner-Elements from its official website:  
[https://karabiner-elements.pqrs.org/](https://karabiner-elements.pqrs.org/)  
2. **Install** the software on your Mac.
3. Open the Karabiner-Elements application.
4. macOS may request accessibility permissions for Karabiner-Elements to function correctly. Grant the necessary permissions as prompted. You might need to reboot or restart the application.

## Configuring the shortcuts

1. Download the default preset keyboard shortcuts:  
[InstrumentaDefault.json](https://github.com/iappyx/Instrumenta-Keys/raw/main/bin/shared-shortcuts/mac/Karabiner-Elements/InstrumentaDefault.json)
2. Open the JSON file in your preferred text editor.
3. Select all content and copy it to your clipboard.
4. Open **Karabiner-Elements Settings**, then navigate to **"Complex Modifications"**.
5. Click **"Add your own rule"**.
6. Select and delete the template JSON, then paste the contents of **InstrumentaDefault.json**.
7. Click **"Save"**.
8. *(Optional)* Navigate to **"Devices"** and enable **"Modify events"** for any external keyboards you may have.
9. Open **PowerPoint** and enjoy your keyboard shortcuts!

## Customize or add Shortcuts
To customize or add shortcuts, you'll need to modify the JSON file used by Karabiner-Elements. The relevant section in the JSON where the script is executed follows this format:
**"/usr/bin/osascript ~/Library/Application\\ Scripts/com.microsoft.Powerpoint/InstrumentaKeys.applescript <Instrumenta Macro Reference>"**

This command requires a **Macro Reference** as a command-line argument. Instrumenta provides various built-in macro references that trigger specific functions. **InstrumentaDefault.json** contains the following shortcuts:

| Shortcut                      | Instrumenta Command                  | Description                                   |
|--------------------------------|--------------------------------------|-----------------------------------------------|
| Ctrl + Shift + S               | ObjectsSwapPosition                 | Swap positions of selected objects.          |
| Ctrl + Shift + L               | ObjectsAlignLefts                   | Align selected objects to the left.          |
| Ctrl + Shift + T               | ObjectsAlignTops                    | Align selected objects to the top.           |
| Ctrl + Shift + R               | ObjectsAlignRights                  | Align selected objects to the right.         |
| Ctrl + Shift + B               | ObjectsAlignBottoms                 | Align selected objects to the bottom.        |
| Ctrl + Shift + E               | ObjectsAlignCenters                 | Align selected objects by their center.      |
| Ctrl + Shift + M               | ObjectsAlignMiddles                 | Align selected objects by their middle.      |
| Ctrl + Shift + H               | ObjectsDistributeHorizontally       | Distribute selected objects horizontally.    |
| Ctrl + Shift + V               | ObjectsDistributeVertically         | Distribute selected objects vertically.      |
| Ctrl + Shift + Left Arrow      | MoveTableColumnLeft                 | Move the selected table column to the left.  |
| Ctrl + Shift + Right Arrow     | MoveTableColumnRight                | Move the selected table column to the right. |
| Ctrl + Shift + Up Arrow        | MoveTableRowUp                      | Move the selected table row up.              |
| Ctrl + Shift + Down Arrow      | MoveTableRowDown                    | Move the selected table row down.            |
| Ctrl + Shift + Q               | GenerateStickyNote                  | Create a sticky note in the presentation.    |

The full list of available macro references can be found [here](https://github.com/iappyx/Instrumenta-Keys/blob/main/instrumenta-macro-list.md). 
Updating the JSON file with your desired shortcuts will allow you to tailor Instrumenta Keys to your workflow.

## How to Build from Source
Building from source is straightforward and allows you to customize Instrumenta Keys to your needs.

### InstrumentaKeysHelper
1. Open **"InstrumentaKeysHelper.pptm"** from the `/src/mac/` directory in PowerPoint.
2. Enable the **Developer** tab in the PowerPoint ribbon through PowerPoint settings.
3. All coding is done in the **Visual Basic Editor (VBA IDE)** of PowerPoint. The `.bas` files in the **Modules** directory are for reference only and are exported after every build.
4. You can modify the `.pptm` file, update existing code, or create your own and copy-paste the relevant sections.
5. To customize the **PowerPoint Ribbon**, use [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor) to edit the `.pptm` file.
6. In PowerPoint, save the file as a **PowerPoint Add-in (`*.ppam`)** to generate your own build.

### InstrumentaKeys.applescript
1. Open **"InstrumentaKeys.applescript"** in your preferred text editor.
2. Modify the script as needed to align with your specific needs.

# Feature requests and contributions
I am happy to receive feature requests and code contributions! Let's make the best toolbar together. For feature requests please create new issue and label it as an enhancement (https://github.com/iappyx/Instrumenta/issues/new/choose). 

- If you want to contribute, please make sure that the code can be freely used as open source code. Please only update the files in /src/. For security reasons I will not accept updated files in /bin/.
- If you want to share your shortcut csv (Windows) or json (Mac), please add them to /shared-shortcuts/
- If you like this Instrumenta & Instrumenta Keys, please let me and the community know how you are using this in your daily work: https://github.com/iappyx/Instrumenta/discussions/5
