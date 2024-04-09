# CLICKSEND VBS

This script processes text files in a specified directory and sends their content to the configured ClickSend API endpoint for SMS messaging.


## Description

The ClickSend.vbs Processor script is designed to automate the process of sending SMS messages via the ClickSend API. It reads text files from a designated folder, extracts the content, and sends it as an SMS message using ClickSend's API. 

## Features

- Reads text files from a specified directory
- Sends SMS messages via ClickSend API
- Supports both .txt and .log file formats
  
## Prerequisites

- Windows operating system
- ClickSend account with API credentials
- Text files containing SMS content in the specified directory

## Installation

1. Download the ClickSend.vbs script from the repository; save it into a the desired location. Recommend a local disk, not a network resource, as VBS files are not trusted on Intranet/Internet Zones following IE Trust Zones default configuration.
2. Ensure that VBScript is enabled on your Windows machine (this could be prevented with Group Policies, or Intune Policies)
3. Create a `Queue` sub-folder and configure whatever product you need to save properly formatted JSON files into that folder. [This document should help](https://developers.clicksend.com/docs/rest/v3/#send-sms).
4. Update the `ClickSend.ini` file with your ClickSend API credentials (API_USERNAME and API_KEY), from your [subaccounts](https://dashboard.clicksend.com/account/subaccounts) page.
5. Place the script in a directory of your choice.
6. Ensure that the directory path for text files and the ClickSend API endpoint are correctly configured in the script.

## Usage

1. Place text files containing SMS content in the designated directory.
2. Run the script by double-clicking on it or executing it via the command line.
3. The script will process the files in the directory and send SMS messages via the ClickSend API.
4. Check the script output for any errors or status messages.

## Configuration

- `ClickSend.ini`: File contains ClickSend API credentials (API_BASE, API_USERNAME, API_KEY).
  
  You will have to updat 
  Also if ClickSend update their API above v3 you will need to update this.
  
- `Queue` sub-directory: Contains text files with SMS content to be processed. However you fill this directory remains out of scope for this product.


## Running this script (TESTING)

Launch a command line, change directory to the location of the script; and execute the following command.
```cmd
cd C:\ClickSend
cscript -nologo ClickSend.vbs
```

## Running this script (SCHEDULED)
1. **Open Task Scheduler:**
   - Press `Win + R` to open the Run dialog.
   - Type `taskschd.msc` and press **Enter**.

2. **Create a New Task:**
   - In the Task Scheduler window, click on "Create Basic Task" in the Actions pane on the right.

3. **Set Task Name and Description:**
   - Enter a name and description for your task, then Click **Next**.

4. **Set Trigger:**
   - Choose "At Startup" as the trigger type and Click **Next**.

5. **Advanced Settings:**
   
   Under the Advanced Settings of the event trigger
   - Enable "Repeat task every", and set to `5 minutes` from the dropdown, for a duration of `5 minutes`. *Consider dropping this to `1 minute` after several days of successful testing.*
   - Click **Next**.

6. **Set Start Date and Time:**
   - Choose the start date and time for the task.
   - Ensure that "Start a program" is selected, then Click **Next**.

7. **Choose Program/Script:**
   - Click Browse and navigate to the location of your script file (e.g., `ClickSend.vbs`).  
     **Note:** Browsing for the file, prevents file-path issues related to spaces or special characters.

   - Select the script file and Click **Next**.

1. **Finish:**
   - Review your task settings and click Finish to create the task.

Once you've completed these steps, the script will be scheduled to run every 5 minutes according to the specified trigger settings. Make sure your computer is powered on and not in sleep mode during the scheduled execution times for the script to run successfully.


## Troubleshooting

- If the script encounters any errors, check the console output for error messages.
- Ensure that the `ClickSend.ini` file is correctly configured with valid ClickSend API credentials.
- Check that the script has permission to read files from the designated directory and send HTTP requests to the ClickSend API.

## License

This script is released under the MIT License. See [License.txt](License.txt) for more information.

## API Documentation
- ClickSend API documentation: [https://developers.clicksend.com/docs/rest/v3](https://developers.clicksend.com/docs/rest/v3)