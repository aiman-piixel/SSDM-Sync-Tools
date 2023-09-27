# SSDM-Sync-Tools

[SSDM-Sync-Tools] - Briefly describe what your project is about.

## Table of Contents

- [Project Overview](#project-overview)
- [Getting Started](#getting-started)
  - [Preparation/Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## Project Overview

Provide a concise overview of your project. Explain its purpose, goals, and any unique aspects that set it apart.

## Getting Started

Provide instructions for users to get your project up and running locally.

### Installation

1. **Download and Install Anaconda**:

   - Download the Anaconda Distribution installer from [Anaconda's official website](https://www.anaconda.com/products/distribution).
   - Run the installer and follow the installation instructions.
   
   During the installation:
   
   - Check the option to "Add Anaconda3 to my PATH environment variable." This ensures that Anaconda's commands are available in your system's command prompt or terminal.
   - Check the option to "Register Anaconda3 as the default Python *.*." This makes Anaconda3 the default Python interpreter.

2. **Create a New Conda Environment**:

   Open a terminal or Anaconda prompt and run the following command to create a new environment named 'py' with the latest version of Python:

   ```sh
   conda create --name py python
   ```

   Then run the following command to activate the new environment and install these packages:

   ```sh
   conda activate py
   conda install pandas jupyter notebook
   ```
   These lines of command will install Pandas, Jupyter Notebook and their dependencies inside the 'py' environment

   While the 'py' environment is active, install additional packages using pip3:

   ```sh
   pip3 install openpyxl selenium glob2 textdistance
   ```
   These commands will install Openpyxl, Selenium, and Glob2 in the 'py' environment.

   To start the Jupyter Notebook server, run the following command:
   ```sh
   jupyter notebook
   ```
   *Note: Some older conda version require this command instead. Try this if the command above does not work and vice versa!
   ```sh
   conda jupyter notebook
   ```

4. **Installing Chrome WebDriver**:

   To use Selenium with Google Chrome, you'll need the Chrome WebDriver. Follow these steps to download and set it up:

  1. **Identify Your Chrome Version**:

       To download the correct version of Chrome WebDriver, you need to know your Chrome browser's version:
   
       - Open Google Chrome.
       - Click on the three vertical dots in the top right corner to open the menu.
       - Go to "Help" > "About Google Chrome".
       - Note the version number displayed. It will be something like `xx.x.xxx.x`.
   
  2. **Downloading Chrome WebDriver**:
      
       If you are using Chrome version 115 or newer, please consult the [Chrome for Testing availability dashboard](https://googlechromelabs.github.io/chrome-for-testing/). This page provides convenient JSON endpoints for specific ChromeDriver version downloading. Look for Stable 64-bit version. Save the downloaded ZIP file to your computer.
      
       If you are using Chrome version 114 or older. Visit the [Chrome WebDriver Download Page](https://sites.google.com/chromium.org/driver/downloads?authuser=0). Look for Stable 64-bit version. Save the downloaded ZIP file to your computer.
      
  3. **Extract the WebDriver Executable**:

   - Locate the downloaded ZIP file and extract its contents to a location on your computer (e.g., `C:\WebDriver`).
   
  4. **Add WebDriver Directory to System PATH**:
  
     - Search for "Environment Variables" in the Windows Start menu and select "Edit the system environment variables".
     - Click the "Environment Variables" button.
     - Under "System variables", find the "Path" variable and click "Edit".
     - Click "New" and enter the full path to the directory containing the WebDriver executable (e.g., `C:\WebDriver`).
     - Click "OK" to close the windows.
     
  5. **Restart Your Computer**:
  
     For the changes to take effect, restart your computer.
  
  6. **Verify WebDriver Installation**:
  
     - Open a Command Prompt.
     - Type `chromedriver` and press Enter. If you see output without errors, the WebDriver is installed correctly.
  
  Now you have the Chrome WebDriver installed and set up on your Windows system. You can use it with Selenium for automated testing and web scraping.

## Usage

  There are 3 main steps to sync SSDM data from StudentQR to SSDM website. Either way, make sure the to launch the appropiate conda environment and jupyter notebook first, then head to this project's folder directory, whenever you placed them.:
  
  - [Downloading Submission](#downloading-submission)
  - [Data Cleaning & Prep](#data-cleaning-&-prep)
  - [Script Preparation](#script-preparation)

  Each steps involves different scripts for modularity, so please make to read carefully on which script to use.

### Downloading Submission

  We must first download and gather all submissions that we would like to sync from the StudentQR admin website. In this steps, you have 2 options. You can either download submissions from one school at a time, or you can download submissions from a list of schools in batch. 

  To download from one school only, open this notebook: **download_ssdm_single_school.ipynb**
  To make it usable on your machine, there are some variables that you need to alter, which are the folder directory, the school email and password if needed, and the calender cycle loop block (more on that later)
   
