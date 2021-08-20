# Download meeting recordings
Downloads webex or zoom recordings onto user provided paths.

## Usage:
> python download_meeting_recordings.py --source <\path\to\source_file.csv> --driver <\path\to\chromedriver.exe> --destination <\path\to\downloads_dir>
- If destination is not provided, uses current directory for downloads.
- If driver is not provided, tries to search for chromedriver in current directory.
- Use --no-prompt/-n to avoid prompting user.

## Pre-requisites:
- Download relevant ChromeDriver from https://chromedriver.chromium.org/downloads place it under utilities folder.
- Install selenium using "pip install selenium".
- Use python version 3.6+ on Windows.

## More Info:
- If browser prompts from CAPTCHA (or) fails with invalid password:
    - When no-prompt mode is disabled: console will ask for user confirmation to proceed further after manually entering captcha.
    - When no-prompt mode is enabled: script execution will fail downloading for that particukar url and proceed further.
- If there is any exception while downloading a link, console shows that there is error downloading that particular link
and proceeds with other urls. At the end, all failed links will be displayed.
