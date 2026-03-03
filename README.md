<p align="center">
  <h1 align="center">Print Spool Optimizer</h1>
  <p align="center">OpenSource Project | Open for Contribution  </p>
  <p align="center">
    A high-performance PDF flattening and compression utility  
    designed to prevent printer memory overflow and optimize print spooling.
  </p>
</p>

<p align="center">
  <img src="assets/pdf-spool-opti-logo.png" alt="Print Spool Optimizer Logo" width="80%">
</p>

<p align="center">
  <a href="#">Python 3.8+</a> •
  <a href="#">MIT License</a> •
  <a href="#">Open Source</a>
</p>

---

## Overview
The Print Spool Optimizer is a Python utility designed to flatten and compress PDF documents. By converting complex vector graphics, fonts, and layers into standard grayscale images, it reduces the processing load on printer hardware. This prevents memory overflow and ensures rapid print spooling for large documents.

<p align="center">
  <img src="assets/UI-Web.png" alt="Web UI" width="80%">
</p>
## Prerequisites
* Python 3.8 or higher

## Installation

To prevent dependency conflicts with your system's global Python installation, you must run this tool within an isolated virtual environment.

### 1. Create a Virtual Environment
Navigate to the root directory of this project in your terminal and execute the following command to create a virtual environment named `venv`:

```bash
python3 -m venv venv
```

### 2. Activate the Virtual Environment
You must activate the environment before installing dependencies or executing the script. Your terminal prompt will change to indicate the environment is active.

#### For macOS / Linux:
```
source venv/bin/activate
```

#### Windows (Command Prompt):
```
venv\Scripts\activate.bat
```

#### Windows (PowerShell):
```
venv\Scripts\Activate.ps1
```

### 3. Install Dependencies

Once the virtual environment is active, install the required packages using the provided requirements file:
```
pip install -r requirements.txt
```

## How to Execute this command :
```
python spool_optimizer.py -i input_file_name.pdf -o output_file_name.pdf
```

### Deactivation
When you are finished using the tool, you can exit the virtual environment by running:
```
deactivate
```

**Author:** Mathi Yuvarajan T.K
