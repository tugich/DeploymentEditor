# DeploymentEditor
Visual Software Packaging Editor by TUGI.CH

<!-- ABOUT THE PROJECT -->
## 🧑‍💻 About The Project
**Welcome to the Deployment Editor for PSADT (PSAppDeployToolkit).** This application simplifies the packaging process for all Windows engineers. You can click your sequence for PSADT through the GUI and with just one simple click you get the fully automated coded PowerShell script for the deployment. The best part: It’s free & open-source.<br/><br/>
![App Screenshot](Screenshot.png)

<!-- GETTING STARTED -->
## ✅ Getting Started

### Requirements
Windows 10/11 64-bit (It also runs well on ARM-based Windows via the translation layer.)<br/>
PureBasic IDE (to compile the source code if needed, binary is included in each release).

### Prerequisites
There are no prerequisites for this tool. All required files are included in this GitHub repository.

### Installation
You can download the latest release from GitHub or use the MSI installer from the releases.<br/>
The editor runs in user context and doesn't require administrator rights. For packaging, you need administrator rights to run a packaging project (in system context).

<!-- TEMPLATE SYSTEM -->
## ✅ (v1.0.7) Project-based templates
There are several templates for each part of the editor. You can customize them and the compiler will take them for each time in the deployment file generation process.
In the latest version (v1.0.7), each project/PSADT package contains its own template (Invoke-AppDeployToolkit.ps1.template) for the final resulting file (Invoke-AppDeployToolkit.ps1).

<!-- FOLDERS AND FILES -->
## 📄 Folders and files
- ***Databases [Folder]:*** PSADT.sql database for defining all commands and parameters.
- ***Examples [Folder]:*** Some basic examples with other software than 7-Zip
- ***Forms [Folder]:*** All forms created with the PureBasic IDE
- ***Plugins [Folder]:*** Built-in editor plugins written in PowerShell
- ***Resources [Folder]:*** Images, icons and more
- ***Scripts [Folder]:*** Some scripts for the development part of the editor
- ***Snippets [Folder]:*** (Not used yet)
- ***Templates [Folder]:*** All templates used by the editor to build any script or file for the final package
- ***Test [Folder]:*** Basic example with 7-Zip installer - ThirdParty: Some third party libraries like PSADT and more
- ***DeploymentEditor.pb [File]:*** The main source file
- ***DeploymentEditor.pbp [File]:*** The project file for the PureBasic IDE

<!-- USAGE EXAMPLES -->
## 📋 Usage

**Video Tutorial on YouTube:**<br/>
[Deployment Editor - Package Softwares with PSAppDeployToolkit (PSADT)](https://www.youtube.com/watch?v=1Ct5B27BGP4)<br/>
There is also an example that shows a simple sequence for installing and uninstalling 7-Zip. Give it a try and if you have any questions just contact me via email or LinkedIn.

<!-- COMPILING -->
## ⚙️ Compile for Windows
You need the PureBasic IDE in the latest version to compile the source code for Windows: https://www.purebasic.com.
Just open the DeploymentEditor.pbp file and run the compiler with [F5] - the rest is magic.

<!-- LICENSE -->
## 📄 License
See `LICENSE.txt` for more information.

<!-- CREDITS -->
## 📄 Credits
[PSAppDeployToolkit/PSAppDeployToolkit](https://github.com/PSAppDeployToolkit/PSAppDeployToolkit)<br/>
See also LICENSE_ThirdParty.txt

<!-- CONTACT -->
## 📧 Contact
TUGI - [contact@tugi.ch](mailto:contact@tugi.ch)<br/>
Project Link: [https://blog.tugi.ch/deployment-editor-preview](https://blog.tugi.ch/deployment-editor-preview)
