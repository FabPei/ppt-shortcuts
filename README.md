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
Instrumenta Keys currently supports **Windows only**.  
An **AppleScript version** is under exploration.

## Installation
Instrumenta Keys is a **PowerShell script** and does not require administrative rights for installation on most enterprise systems.

### Windows Installation
1. **Download the binary:** [Instrumenta Keys.exe](https://github.com/iappyx/Instrumenta-Keys/raw/main/bin/windows/Instrumenta%20Keys.exe)
2. Run the binary. It will automatically **generate a shortcut file for you** and **minimize to the system tray** after a few seconds. Enjoy your shortcuts!
3. To open it again, **click the Instrumenta Keys icon** in the system tray. If you click the icon again it will hide again.

Note: A security notice may appear when running the binary, please refer to [these](https://support.microsoft.com/en-gb/topic/a-potentially-dangerous-macro-has-been-blocked-0952faa0-37e7-4316-b61d-5b5ed6024216) instructions from Microsoft to unblock Instrumenta Keys: (1) Open Windows File Explorer and go to the folder where you saved the file; (2) Right-click the file and choose Properties from the context menu; (3) At the bottom of the General tab, select the Unblock checkbox and select OK.

### How to load shortcuts
Shortcuts in Instrumenta Keys are stored in the `shortcuts.csv` file. If this file is missing when you launch Instrumenta Keys, a default version will be created automatically. You can manage and customize your shortcuts using the Shortcut Editor, which you can open in PowerPoint by pressing CTRL+SHIFT+ALT+Q. This editor allows you to modify your shortcuts, import or export shortcut files, and even load shortcut configurations directly from this GitHub page. This makes it easy to use shortcut files shared by others, such as sets tailored to specific consultancy firm standards.
Instructions and a full list of available macros in Instrumenta can be found [here](https://github.com/iappyx/Instrumenta-Keys/blob/main/instrumenta_macros.md)

![image](https://github.com/user-attachments/assets/449c14ab-799e-4377-a249-f318118baddb)

## How to Build from Source
Building your own Instrumenta Keys is very simple:

### Requirements
- **Microsoft PowerShell**
- **PS2EXE module** (PowerShell-to-EXE converter)

### Steps
1. Locate the source code in `\src\src.ps1`
2. Run the build script in an elevated PowerShell window: `.\build.ps1`
3. The executable will be generated in `\bin\windows\Instrumenta Keys.exe`

# Feature requests and contributions
I am happy to receive feature requests and code contributions! Let's make the best toolbar together. For feature requests please create new issue and label it as an enhancement (https://github.com/iappyx/Instrumenta/issues/new/choose). 

- If you want to contribute, please make sure that the code can be freely used as open source code. Please only update the files in /src/. For security reasons I will not accept updated .exe files.
- If you want to share your shortcut csv-file, please add it to /shared-shortcuts/
- If you like this Instrumenta & Instrumenta Keys, please let me and the community know how you are using this in your daily work: https://github.com/iappyx/Instrumenta/discussions/5
