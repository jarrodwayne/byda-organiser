# Import necessary modules to enable program functionality
import os
import sys
import re
import io
import shutil
import win32com.client
from tkinter import filedialog
from pathlib import Path
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
import configparser
import glob
import time
import datetime
import pytz
from plyer import notification as plyer_notification
import pystray
from PIL import Image
import threading
import requests

# Retrieve the directory where the Python script is located
script_dir = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))

# Specify the relative path to the program logo based on the script_dir
image_path = os.path.join(script_dir, "foreman_software_solutions_logo.png")

# Specify the relative path to the program logo based on the script_dir
icon_path = os.path.join(script_dir, "foreman_software_solutions_logo.ico")

# Create a global variable to track the status of the system tray icon
system_tray_icon_status = False


# Check if the user's OS has an active internet connection
def check_internet_connection():
    # Attempt to reach the specified web address below
    try:
        print("Connecting to Online Services...")
        # If the program is able to establish a connection, return a True value
        requests.get("http://www.google.com.au", timeout=5)
        print("Connected to Online Services Successfully.")
        return True
    # If the program cannot establish a connection, return a False value
    except requests.ConnectionError:
        print("Failed to Connect to Online Services. Please Ensure An Active Connection is Available.")
        return False


# Create a system tray icon that runs alongside the main program function
def initialize_system_tray_icon():
    # Enable exit functionality within the system tray icon
    def on_exit(tray_icon):
        global system_tray_icon_status
        icon.visible = False
        tray_icon.stop()
        system_tray_icon_status = False
        return

    print("Initializing System Tray Icon...")
    print("System Tray Icon Created.")

    try:
        # Create the system tray icon menu
        menu = (pystray.MenuItem('Stop', on_exit),)
        # Establish the system tray icon image and icons
        image = Image.open(image_path)
        icon = pystray.Icon("name", image, "BYDA Organiser", menu)
        icon.run()
        icon.visible = True

    except AttributeError:
        print(f"Error: Unable to Initialize System Tray Icon.")
        return


# Initialize the processed_jobs set from the configuration file
def initialize_config_file():
    try:
        print("Initializing Configuration File...")

        # Establish config file parameters
        config = configparser.ConfigParser()
        # Determine the config file pathway/name
        config_file_path = "config.ini"
        # Read the config file
        config.read(config_file_path)

        load_processed_jobs = config.get("Statistics", "processed_jobs", fallback="")
        # Load the processed_jobs set from the config file
        processed_jobs = set()

        if load_processed_jobs:
            processed_jobs = set(int(job) for job in load_processed_jobs.split(',') if job.isdigit())

        # Load the job_count integer from the config file
        job_count = config.getint('Statistics', 'job_count', fallback=len(processed_jobs))

        print("Configuration File Initialized.")

    except AttributeError:
        print(f"Error: Configuration File Could Not Be Initialized.")
        return

    # Return the processed_jobs set to be accessed in other functions
    return processed_jobs, job_count


# Request the target directory input and Microsoft Outlook folder from the program user
def retrieve_user_input():
    # Temporarily initialize the Microsoft Outlook application to allow for user input
    outlook = win32com.client.DispatchEx("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")

    try:
        print("Awaiting User Input...")

        # Get the target Microsoft Outlook folder input from the user
        target_inbox_folder = namespace.PickFolder()

        # Check if the user closed the Microsoft Outlook folder selection screen
        # Return a 'None' value if the user closes the respective screen
        if not target_inbox_folder:
            return None, None
        else:
            print(f"Microsoft Outlook Folder Input Retrieved.")

        # Get the target directory input from the user
        target_directory_input = filedialog.askdirectory(title=f"Select A Target Directory")

        # Check if the user closed the target directory selection screen
        # Return a 'None' value if the user closes the respective screen
        if not target_directory_input:
            return None, None
        else:
            print(f"Target Directory Input Retrieved.")

        # Convert the target directory output location variable to string
        user_output_location_name = str(target_directory_input)

        # Define the user selected output directory as a pathway
        target_directory = Path(user_output_location_name.strip())

        print("User Input Retrieved.")

    except AttributeError:
        print(f"Error: Unable to Retrieve User Inputs.")
        return

    # Return the user input variables to be used in other functions
    return target_directory, target_inbox_folder


# Initialize Microsoft Outlook application
def initialize_outlook(target_inbox_folder):
    try:
        # Connect to Microsoft Outlook application
        print("Connecting to Microsoft Outlook Application...")
        outlook = win32com.client.DispatchEx("Outlook.Application")
        outlook.GetNamespace("MAPI")

    except AttributeError:
        print(f"Error: Unable to Connect to Microsoft Outlook Application.")
        return

    print("Connected to Microsoft Outlook Application.")

    try:
        # Locate the Microsoft Outlook Application Inbox folder
        print("Locating Inbox Folder...")
        inbox = target_inbox_folder

    except AttributeError:
        print(f"Error: Unable to Locate/Access The Specified Mailbox Folder.")
        return

    print("Inbox Folder Identified.")

    # Return the inbox variable to be used in future functions
    return inbox


# Retrieve BYDA job numbers from the user-selected Microsoft Outlook folder
def retrieve_job_information(inbox, processed_jobs, target_directory):
    # Create an empty list variable to store the job numbers
    job_numbers = []

    # Determine the timezone to be used within the scanning criteria
    scanning_time_criteria_timezone = pytz.timezone('Australia/Melbourne')
    # Calculate the email message scanning time criteria (14 days/2 weeks)
    scanning_time_criteria = datetime.datetime.now(scanning_time_criteria_timezone) - datetime.timedelta(days=14)

    print(f"Retrieving Job Information...")

    # Iterate over the email messages within the user-selected Microsoft Outlook folder
    for email in inbox.Items:
        # Match the email message files to the relevant criteria and check if it has been received within the
        # scanning_time_criteria timeframe
        if email.SenderName.lower() == "dbyd@1100.com.au" and email.ReceivedTime >= scanning_time_criteria:
            # Extract job numbers from the email message subject lines using the specified criteria
            subject_criteria = email.Subject.lower()
            message_job_number = re.findall(r'\b\d{8}\b', subject_criteria)
            # If any 8-digit numbers are found in the subject line, add them to the job_numbers list
            job_numbers.extend(message_job_number)
            print(f"Job {message_job_number} Information Identified.")

    # Do not remove contents of target directory if job_number is part of processed_jobs set in config file
    for job_number in job_numbers:
        for item in target_directory.iterdir():
            if (str(job_number) in str(target_directory) and target_directory.is_dir() and job_number not in
                    processed_jobs):
                if item.is_file():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
            else:
                # Create the target directory if it doesn't exist already
                target_directory.mkdir(parents=True)

    # Return specific variables to be used in other functions
    return job_numbers, scanning_time_criteria


# Create BYDA job number subdirectories from the corresponding identified email message files
def initialize_byda_job(job_number, target_directory, scanning_time_criteria, inbox):
    # Initialize module variables and values
    invalid_filename_characters = r'[<>:"/\\|?*]'
    byda_job_location = None

    print(f"IMPORTANT: Processing Job {job_number}...")
    print(f"Creating Service Provider Subdirectories...")

    # Iterate over the email messages within the user-selected Microsoft Outlook folder
    for email in inbox.Items:
        try:
            # Normalize the coversheet message E-Mail address and define message subject parameters
            normalized_sender_address = email.SenderEmailAddress.lower().strip().replace('<', '').replace('>', '')
            sender_subject = email.Subject.strip()

            # Check if the email meets the time and sender criteria for relevant job identification
            if email.ReceivedTime >= scanning_time_criteria and normalized_sender_address == "dbyd@1100.com.au":
                # Use a regular expression to strictly match the subject line format of the provider message
                coversheet_message_format = re.compile(r'^BYDA JOB: (\d+) - .+', re.IGNORECASE)
                coversheet_message_match = coversheet_message_format.match(sender_subject)

                # If the subject line matches the expected format, determine the relevant variables for processing
                if coversheet_message_match and str(job_number) in sender_subject:
                    # Determine and create the BYDA job subdirectory name and pathway
                    byda_job_name = re.sub(invalid_filename_characters, '', sender_subject)
                    byda_job_location = Path(target_directory / byda_job_name.strip())
                    byda_job_location.mkdir(parents=True, exist_ok=True)
                    print(f"{byda_job_location}")

        except AttributeError:
            print(f"Warning: Failed to Create Subdirectory for Job {job_number}.")
            continue

    print(f"Provider Subdirectories Created.")

    # Return the necessary variables and values to be accessed in other functions
    return byda_job_location


# Make a copy of the provider email message files and move them into the 'E-Mail Files' subdirectory
def copy_message_files(job_number, byda_job_location, outlook, scanning_time_criteria, inbox):
    # Initialize module variable
    invalid_filename_characters = r'[<>:"/\\|?*]'

    print(f"Copying Provider Message Files...")
    # Scan the email messages within the user's selected folder in Microsoft Outlook
    for email in inbox.Items:
        try:
            # Check if the job number being processed matches the criteria outlined below
            if str(job_number) in email.Subject.lower() and email.ReceivedTime >= scanning_time_criteria:
                # Create the message filename and assign its relevant location
                provider_message_filename = str(re.sub(invalid_filename_characters, '', email.Subject
                                                       + '.msg'))
                provider_message_location = Path(byda_job_location / "E-Mail Files".strip())
                provider_message_location.mkdir(parents=True, exist_ok=True)
                try:
                    # Save the email message files to the predetermined destination
                    provider_message_file = outlook.Session.GetItemFromID(email.EntryID)
                    provider_message_file.SaveAs(str(provider_message_location / provider_message_filename))

                except AttributeError:
                    print(f"Warning: Failed to Save Message File {provider_message_filename}.")
                    continue

        except AttributeError:
            print(f"Error: Unable to Copy Job Message Files.")
            return

    print(f"Provider Message Files Copied.")


# Extract the attachments from the provider email messages and move them into their previously-created subdirectories
def extract_message_files(job_number, byda_job_location, scanning_time_criteria, inbox):
    # Initialize module variables
    provider_coversheet_name = ''
    invalid_filename_characters = r'[<>:"/\\|?*]'
    kdr_indicator = False

    print(f"Extracting Provider Plans...")
    # Loop through the email messages within the user's selected folder in Microsoft Outlook
    for email in inbox.Items:
        try:
            # Check if the job number being processed matches the criteria outlined below
            if str(job_number) in email.Subject.lower() and email.ReceivedTime >= scanning_time_criteria:
                # Determine the provider subdirectory name given the specified criteria
                provider_name = str(email.SenderName.replace("BYDA -", "").strip() or "Unknown Provider")

                # Check if the sender's email address matches the specific address
                if provider_name == "dbyd.JENreplyTA@jemena.com.au":
                    provider_name = "Jemena Electricity Networks (VIC)"

                # Flag the kdr_indicator as True if the specific provider name is identified
                if provider_name == "KDR Victoria Pty Ltd":
                    kdr_indicator = True

                provider_location = Path(byda_job_location / provider_name.strip())
                provider_location.mkdir(parents=True, exist_ok=True)
                try:
                    # Save the attachments within their respective subdirectories
                    for attachment in email.Attachments:
                        attachment_filename = str(re.sub(invalid_filename_characters, '', attachment.FileName))
                        attachment_location = os.path.join(provider_location, attachment_filename.strip())
                        attachment.SaveAsFile(attachment_location)

                except AttributeError:
                    print(f"Warning: Failed to Extract {provider_name} Message Attachments.")
                    continue

        except AttributeError:
            print(f"Error: Failed to Extract Job Message Files.")
            return

    print(f"Provider Plans Extracted Successfully.")

    # Return the necessary variable to be accessed in other functions
    return provider_coversheet_name, kdr_indicator


# Identify the job coversheet file in the BYDA job location
def initialize_coversheet(job_number, byda_job_location):
    print(f"Locating Coversheet File...")

    # Determine the coversheet file, the new filename, and the new location
    coversheet_file = f"{job_number}.pdf"
    coversheet_filename = f"Job {job_number} - Cover Sheet.pdf"
    coversheet_location = os.path.join(byda_job_location, "dbyd@1100.com.au", coversheet_file.strip())

    # Check if the expected coversheet file exists within the BYDA subdirectory
    if os.path.isfile(coversheet_location):
        try:
            # Rename the coversheet file and move it to the BYDA job location directory
            shutil.move(coversheet_location, os.path.join(byda_job_location, coversheet_filename))
            # Remove the subdirectory where the coversheet file was originally located
            shutil.rmtree(os.path.join(byda_job_location, "dbyd@1100.com.au").strip())

        except AttributeError:
            print(f"Error: Failed to Locate Job Coversheet File.")
            return

    print(f"Coversheet File Located.")


# Scan the job coversheet file to identify if any provider plans are missing
def scan_coversheet(job_number, byda_job_location):
    # Define/initialize variables to be used later in the function
    coversheet_extracted_text = []
    end_index = 0

    print(f"Analysing Coversheet File...")

    # Define specific-use-case variables to be utilized in the loop
    coversheet_location = os.path.join(byda_job_location, f"Job {job_number} - Cover Sheet.pdf".strip())
    coversheet_manager = PDFResourceManager()
    try:
        # Open the coversheet file and extract the text contents within
        with open(str(coversheet_location), 'rb') as pdf_file:
            coversheet_string = io.StringIO(initial_value='', newline=None)
            codec_arguments = {}
            la_parameters = LAParams()
            text_converter = TextConverter(coversheet_manager, coversheet_string,
                                           **codec_arguments,
                                           laparams=la_parameters)
            text_interpreter = PDFPageInterpreter(coversheet_manager, text_converter)

            # Iterate over all the pages within the coversheet file
            for page in PDFPage.get_pages(pdf_file):
                text_interpreter.process_page(page)
                coversheet_extracted_text += coversheet_string.getvalue()

            # Define search scope parameters for the scanning process to identify the specific provider details
            coversheet_input_text = coversheet_string.getvalue()
            start_index = coversheet_input_text.lower().find('authority name', end_index)
            end_index = coversheet_input_text.lower().find('END OF UTILITIES LIST', start_index)

            coversheet_output_text = [
                re.sub(r'^\s+|\s+?$', '', re.sub(r'[()\d\']', '', line.strip())).lower()
                for line in coversheet_input_text[start_index:end_index].split('\n')
                if line.strip() and not line.isspace() and line.strip() != ''
            ]

            if "victoria university" in coversheet_output_text:
                coversheet_output_text.remove("victoria university")

            # Retrieve the provider names from the subdirectory names created in the create_subdirectories function
            provider_subdirectory_names = [
                re.sub(r'[()]', '', dir_name.name.lower())
                for dir_name in os.scandir(byda_job_location)
                if dir_name.is_dir() and dir_name.name.lower() != 'e-mail files'
            ]

            print(f"{coversheet_output_text}")

    except AttributeError:
        print(f"Error: Failed to Read Job Coversheet File.")
        return

    print(f"Coversheet Analysis Successful.")

    # Return user variables to be used in other future functions
    return coversheet_output_text, provider_subdirectory_names


# Create a notification panel to display windows notifications for the return_coversheet_results function
def initialize_notification(title, message, app_name="BYDA Organiser"):
    try:
        if plyer_notification:
            # Format the notification window with title, message and app_name variables
            plyer_notification.notify(
                title=title,
                message=message,
                app_name=app_name,
                app_icon=icon_path
            )
        else:
            # If plyer fails to notify, print a notification to the console instead
            print(f"Notification: {title}\n{message}")

    except Exception as e:
        print(f"Error in Notification: {e}")


# Return the scan_coversheet function results to the user via the console and windows notifications
def return_coversheet_results(byda_job_location, kdr_indicator, coversheet_output_text, provider_subdirectory_names,
                              job_number):
    # Initialize module lists and variables
    missing_providers = [provider for provider in coversheet_output_text if provider not in provider_subdirectory_names]
    dwf_indicator = glob.glob(str(byda_job_location / '**/*.dwf').strip(), recursive=True)
    byda_job_complete = len(missing_providers) == 0
    missing_providers_file = None

    print(f"Checking Provider Plans...")

    try:
        # Loop through the extracted text from the coversheet and compare it to the provider subdirectory names
        for text_item in coversheet_output_text:
            text_item_lower = text_item.lower()
            # Check if the text item is not in provider subdirectory names. Mark the job as incomplete if so
            if text_item_lower not in provider_subdirectory_names:
                missing_providers.append(text_item)

            # Create a text file with a list of the missing job providers (if applicable)
            missing_providers_file = os.path.join(byda_job_location, "Missing Providers.txt")
            with open(missing_providers_file, 'w', encoding='utf-8') as file:
                file.write("Missing Providers\n\n")
                for provider_name in missing_providers:
                    formatted_provider_name = provider_name.title()
                    file.write(formatted_provider_name + '\n')

                # Display the missing provider plans to the user as well as other relevant information
            if missing_providers:
                print(f"IMPORTANT: Job {job_number} Processing is Complete.")
                print(f"WARNING: The Following Provider Plans Have Been Identified As Missing:")
                for provider_name in missing_providers:
                    print(provider_name)

        # Return a unique job complete message if all plans have been received
        if byda_job_complete:
            print(f"IMPORTANT: Job {job_number} Processing is Complete.")
            print(f"All Service Provider Plans Have Been Received.")

            # Create the plyer notification window to display to the user
            notification_title = f"NOTICE: Job {job_number}"
            notification_message = f"Processing is Complete. All Service Provider Plans Have Been Received."
            initialize_notification(notification_title, notification_message)

        # Check if all job plans have been received and the "Missing Providers.txt" file exists
        if byda_job_complete and missing_providers_file is not None and os.path.exists(missing_providers_file):
            # Delete the "Missing Providers.txt" file if it exists whilst all plans have been received
            os.remove(missing_providers_file)

        # Inform the user if KDR provider plans have been identified
        if kdr_indicator:
            print(f"IMPORTANT: KDR Victoria Services Have Been Identified.")

            # Create the plyer notification window to display to the user
            notification_title = f"NOTICE: Job {job_number}"
            notification_message = f"KDR Victoria Services Have Been Identified."
            initialize_notification(notification_title, notification_message)

        # Inform the user if .dwf files have been identified
        if dwf_indicator:
            print(f"IMPORTANT: A .dwf Telstra Plan Has Been Identified.")

            # Create the plyer notification window to display to the user
            notification_title = f"NOTICE: Job {job_number}"
            notification_message = f"A .dwf Telstra Plan Has Been Identified."
            initialize_notification(notification_title, notification_message)

    except AttributeError:
        print(f"Warning: Failed to Identify/Confirm Missing Provider Plans.")
        return

    # Return the variable to be used in a future function
    return byda_job_complete


# Create and/or initialize the program configuration file to track key processing statistics
def update_config_file(job_number, byda_job_complete):
    try:
        # Initialize the configuration file
        config = configparser.ConfigParser()
        config_file_path = "config.ini"

        # Load or create the configuration file if necessary
        if not Path(config_file_path).is_file():
            # Create the configuration file with default values if it doesn't already exist
            config["Statistics"] = {"processed_jobs": "", "job_count": "0"}

        # Read the configuration file
        config.read(config_file_path)

        # Load the necessary values from the configuration file
        load_processed_jobs = config.get("Statistics", "processed_jobs", fallback="")
        processed_jobs = set()

        # Format the processed_jobs set
        if load_processed_jobs:
            processed_jobs = set(int(job) for job in load_processed_jobs.split(',') if job.isdigit())

        # Check if the job number has already been processed by the program
        if str(job_number) not in load_processed_jobs and byda_job_complete:
            # Update the set of processed jobs if it hasn't already been processed prior
            processed_jobs.add(job_number)
            # Increment the job count based on the job numbers within the file
            job_count = len(processed_jobs)

            # Update the configuration file with the new processed jobs if necessary
            config.set("Statistics", "processed_jobs", ",".join(map(str, processed_jobs)))
            config.set("Statistics", "job_count", str(job_count))

            with open(config_file_path, "w") as configfile:
                config.write(configfile)

    except AttributeError:
        print(f"Error: Failed to Initialize/Access Configuration File.")
        return


# Define a main function to run the other functions for each job_number
def main():
    # Check for an active internet connection before proceeding with processing
    if not check_internet_connection():
        initialize_notification(f"Cannot Connect to Online Services", f"Please Ensure You Have An Active"
                                                                      "Internet Connection and Try Again.")
        return

    # Initialize variables for the system tray icon
    global system_tray_icon_status
    system_tray_icon_status = True

    # Initialize system tray icon
    icon_thread = threading.Thread(target=initialize_system_tray_icon)
    icon_thread.daemon = True
    icon_thread.start()

    # Step 1: Request input from the program user
    target_directory, target_inbox_folder = retrieve_user_input()

    while system_tray_icon_status:
        # Check for the "Exit" option in the system tray icon
        if not icon_thread.is_alive():
            system_tray_icon_status = False

        # Check if either target_directory or target_inbox_folder has been inputted by the user. End if they're empty
        if target_directory is None or target_inbox_folder is None:
            system_tray_icon_status = False
            break

        # Initialize Microsoft Outlook application
        outlook = win32com.client.DispatchEx("Outlook.Application")
        inbox = target_inbox_folder

        # Step 2: Retrieve BYDA job numbers from the user-selected Outlook inbox folder
        job_numbers, scanning_time_criteria = retrieve_job_information(inbox)

        # Initialize the processed jobs set
        processed_jobs, job_count = initialize_config_file()

        # Check if there are any job numbers to process
        if job_numbers:
            for job_number in job_numbers:
                job_number = int(job_number.strip())
                # Check if the job number has already been processed
                if job_number in processed_jobs:
                    print(f"Note: Job {job_number} Has Already Been Completed. Identifying Next Job Number...")
                    continue

                # Step 3: Create BYDA job subdirectories and copy email message files into designated subdirectories
                byda_job_location = initialize_byda_job(job_number, target_directory, scanning_time_criteria,
                                                        inbox)

                # Step 4: Make a copy of the provider email message files and move them into their subdirectories
                copy_message_files(job_number, byda_job_location, outlook, scanning_time_criteria, inbox)

                # Step 5: Extract the attachments from the provider email messages and move them
                kdr_indicator = extract_message_files(job_number, byda_job_location, scanning_time_criteria, inbox)

                # Step 6: Identify the job coversheet file
                initialize_coversheet(job_number, byda_job_location)

                # Step 7: Scan the job coversheet to identify if any provider plans are missing
                coversheet_output_text, provider_subdirectory_names = scan_coversheet(job_number,
                                                                                      byda_job_location)

                # Step 8: Display the scan_coversheet function results to the user
                byda_job_complete = return_coversheet_results(byda_job_location, kdr_indicator,
                                                              coversheet_output_text,
                                                              provider_subdirectory_names, job_number)

                # Step 9: Create and/or initialize the program configuration file to track key statistics
                update_config_file(job_number, byda_job_complete)

        if system_tray_icon_status:
            print(f"Important: All Current Jobs Have Been Identified/Sorted. Attempting Processing Again In "
                  f"15 Minutes.")
            # Wait for 15 minutes (900 seconds) before running the specific functions again
            # Check for system tray icon user input every 5 seconds
            total_wait_time = 900
            check_interval = 5
            time_elapsed = 0
            # Hibernate the program inbetween specified wait periods
            while time_elapsed < total_wait_time:
                if not system_tray_icon_status:
                    break
                time.sleep(check_interval)
                time_elapsed += check_interval


if __name__ == "__main__":
    main()
