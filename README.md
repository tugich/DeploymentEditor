# DeploymentEditor  
**Visual Software Packaging Editor for PSAppDeployToolkit (PSADT)**  
by **TUGI.CH**

![App Screenshot](Screenshot.png)

---

## 🧑‍💻 About the Project

**DeploymentEditor** is a **visual, project-based software packaging editor** designed to dramatically simplify the creation of professional deployment packages using **PSAppDeployToolkit (PSADT)**.

Instead of manually writing and maintaining complex PowerShell deployment scripts, DeploymentEditor allows Windows engineers to **build installation, uninstallation and repair logic visually** through an intuitive GUI.  
With just a few clicks, the editor **automatically generates a fully structured, production-ready PowerShell deployment script** that follows PSADT best practices.

### Key Benefits
- 🚀 **Accelerated packaging workflow** – build PSADT packages faster and with fewer errors  
- 🧩 **Visual sequence editor** – no deep PowerShell knowledge required  
- 📦 **Project-based architecture** – each package is self-contained and portable  
- 🛠 **Highly customizable** – templates, commands, and plugins are fully editable  
- 🔓 **Free & Open Source** – no licensing costs, no limitations  

DeploymentEditor is ideal for:
- Endpoint Management Engineers (Intune, SCCM, MECM)
- Enterprise IT Administrators
- Software Packaging Specialists
- Anyone deploying Windows software at scale

---

## ✅ Getting Started

### System Requirements
- **Windows 10 / Windows 11 (64-bit)**
- Fully compatible with **ARM-based Windows** via the Microsoft x64 translation layer  
- **PureBasic IDE** (only required if you want to compile the source yourself; precompiled binaries are provided)

### ⚠️ Important: PowerShell Execution Policy (Windows Settings)
DeploymentEditor generates and executes PowerShell scripts based on PSAppDeployToolkit (PSADT).
For this reason, PowerShell script execution must be enabled on the system where packages are built or tested.

### Prerequisites
There are **no external prerequisites**.  
All required components, templates, databases, and resources are included directly in this repository or the release packages.

---

## 📦 Installation

You have two installation options:

1. **Portable / ZIP Version**  
   - Download the latest release from GitHub  
   - Extract and run immediately (no installation required)

2. **MSI Installer**  
   - Available in the GitHub Releases section (if provided)
   - Ideal for managed environments

### Execution Context
- The **editor itself runs in user context** and **does not require administrator privileges**
- **Administrator rights are only required** when executing a packaging project that builds or tests a deployment in **system context**

---

## 🧩 Project-Based Template System (v1.0.7+)

Starting with **version 1.0.7**, DeploymentEditor introduces a **fully project-scoped template system**.

### How It Works
- Each deployment project contains its **own PSADT template**
- Templates are used during compilation to generate the final deployment script
- This allows **maximum flexibility** between different customers, environments, or packaging standards

### Key Template File
- `Invoke-AppDeployToolkit.ps1.template`  
  → Compiled into  
- `Invoke-AppDeployToolkit.ps1`

This approach ensures:
- No global template conflicts  
- Full version control per project  
- Clean separation between logic and presentation  

---

## 📄 Repository Structure

The repository is organized to clearly separate logic, UI, resources, and extensibility components:

- **Databases/**  
  PSADT.sql database defining all available commands, parameters, and metadata

- **Examples/**  
  Sample projects demonstrating packaging workflows beyond 7-Zip

- **Forms/**  
  All GUI forms created using the PureBasic IDE

- **Plugins/**  
  Built-in editor plugins written in PowerShell for extended functionality

- **Resources/**  
  Images, icons, UI assets, and branding resources

- **Scripts/**  
  Helper scripts used during development of the editor

- **Snippets/**  
  Reserved for future expansion (currently unused)

- **Templates/**  
  Core templates used to generate deployment scripts and package files

- **Test/**  
  Example project using the 7-Zip installer  
  Includes third-party components such as PSADT

- **DeploymentEditor.pb**  
  Main PureBasic source file

- **DeploymentEditor.pbp**  
  PureBasic project file

---

## 📋 Usage

### Video Tutorial
📺 **YouTube Walkthrough:**  
[Deployment Editor – Package Software with PSAppDeployToolkit (PSADT)](https://www.youtube.com/watch?v=1Ct5B27BGP4)

The tutorial demonstrates:
- Creating a new project  
- Building install & uninstall sequences  
- Packaging 7-Zip as a real-world example  

If you have questions, feedback, or feature requests, feel free to reach out.

---

## ⚙️ Compiling from Source

To compile DeploymentEditor yourself:

1. Install the **latest version of the PureBasic IDE**  
   https://www.purebasic.com
2. Open `DeploymentEditor.pbp`
3. Compile and run using **F5**

The build process is fully self-contained — no additional dependencies required.

---

## 📄 License

This project is licensed under the terms described in `LICENSE.txt`.

---

## 📄 Credits & Third-Party Software

- PSAppDeployToolkit  
  https://github.com/PSAppDeployToolkit/PSAppDeployToolkit  

Additional third-party licenses can be found in `LICENSE_ThirdParty.txt`.

---

## 📧 Contact

**TUGI**  
📩 Email: [contact@tugi.ch](mailto:contact@tugi.ch)  
🌐 Project Page: https://blog.tugi.ch/deployment-editor-preview  
