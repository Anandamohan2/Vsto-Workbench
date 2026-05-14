# PowerPoint VSTO Add-in

A Microsoft PowerPoint Add-in developed using VSTO (Visual Studio Tools for Office) and C# to demonstrate PowerPoint Interop automation, Ribbon customization, and presentation context analysis.

This project focuses on reading presentation metadata, active slide details, and selected shape information through a custom Ribbon command integrated into PowerPoint.

---

## Features

* Custom Ribbon integration in PowerPoint
* Read active presentation information
* Display total slide count
* Retrieve active slide number
* Read selected shape count
* Display selected shape names and types
* Exception handling and validation support
* PowerPoint Interop Object Model usage

---

## Technology Stack

| Technology           | Description                      |
| -------------------- | -------------------------------- |
| C#                   | Application development language |
| .NET Framework 4.7.2 | Runtime framework                |
| VSTO                 | Office Add-in framework          |
| PowerPoint Interop   | Office automation APIs           |
| WinForms             | UI framework                     |
| Visual Studio 2022   | Development IDE                  |

---

## Implemented Components

### Ribbon Integration

The project includes a custom Ribbon tab and group with a button to read presentation information from the active PowerPoint session.

### Presentation Context Reading

The add-in retrieves:

* Presentation name
* Total slide count
* Active slide index
* Selected shape information

using the PowerPoint Interop Object Model.

### Shape Selection Analysis

The solution validates user selection and iterates through the selected shapes to retrieve:

* Shape Name
* Shape Type

using ShapeRange and Selection APIs.

### Add-in Lifecycle

The add-in startup and shutdown lifecycle are managed through the VSTO host infrastructure.

---

## Learning Objectives

This project was created to strengthen understanding of:

* PowerPoint Object Model
* Office Interop APIs
* Ribbon lifecycle
* VSTO Add-in architecture
* Presentation automation
* Shape and selection handling
* Office event lifecycle
* Enterprise add-in structure

---

## Setup Instructions

### Prerequisites

Install the following:

* Visual Studio 2022 Community/Professional
* VSTO workload
* Microsoft Office PowerPoint
* .NET Framework 4.7.2 Developer Pack

---

### Run the Project

1. Clone the repository

```bash
git clone https://github.com/Anandamohan2/Vsto-Workbench.git
```

2. Open solution in Visual Studio

```txt
PowerPointVstoAdd-In.sln
```

3. Build the solution

4. Press `F5`

5. PowerPoint launches with the custom Ribbon tab

---

## Ribbon Functionality

### Read Presentation Info

The Ribbon button performs:

* Active presentation validation
* Slide count retrieval
* Active slide detection
* Selected shape enumeration
* Presentation diagnostics display

using PowerPoint Interop automation APIs.

---

## Architecture Overview

```txt
PowerPoint
   │
   ├── VSTO Runtime
   │
   ├── ThisAddIn
   │
   ├── Ribbon Layer
   │      └── PresentationInfoRibbon
   │
   └── PowerPoint Interop Object Model
           ├── Application
           ├── Presentation
           ├── Slide
           ├── Selection
           └── ShapeRange
```

---

## Future Enhancements

* Shape creation automation
* Slide export utilities
* Presentation validation engine
* Custom task pane integration
* Chart and table automation
* Logging framework integration
* Ribbon dynamic controls
* Presentation formatting tools

---

## Repository Standards

This repository follows:

* `.gitignore` based source control management
* Clean repository structure
* Separation of generated and source files
* Enterprise-style commit workflow

---

## Author

**Pravin**
Software Developer | VSTO | Office Automation | C# | .NET
