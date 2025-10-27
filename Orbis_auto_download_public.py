
'''
------------------------------------------------------------------------------------------------------------------------
                                                        April 2024
------------------------------------------------------------------------------------------------------------------------



                                            ORBIS DATABASE DOWNLOAD AUTOMATION

                                            

------------------------------------------------------------------------------------------------------------------------

                                                     Gabriele Nespoli
                                                    
------------------------------------------------------------------------------------------------------------------------                                                    
                                                    Bocconi University

                                 You can contact me at:   gabriele.nespoli@studbocconi.it
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------

This code was originally developed by Gabriele Nespoli for data collection purposes in the context of a Bachelor's 
Thesis project at Bocconi University, Milan. It was subsiquently made publicly available to fellow students and 
researchers in order to foster free research, facilitating the data collection process to everyone.

------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------


GENERAL INFORMATION AND TERMS OF USE


By making use of this program, you implicitly accept the following:

This code was developed to automate the task of downloading large amounts of data from the Orbis database by 
Bureau Van Dijk. 
Access is only possible through your institutions' login page. Username and password for the login must be stored in two 
different .txt files, of which you will have to provide the paths.
The only browser currently supported is Safari.
In order for the program to be able to run, you will need to download the Selenium library first.
The XPaths used are based on the website's HTML as of April 2024. You may need to replace some of them if updates occur.

The following code is intended for personal use and research purposes only. Any commercial or different type of use is 
severely forbidden, as it might be against against local law or Orbis' terms of use.
Any abuse is the sole resposibility of the user, and the author does not respond in any way to damages caused to the 
database or to its provider by the misuse of this code.

------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------


INSTRUCTIONS

Before running the code, you will need some set ups according to your research needs.
First, assign values to all the constants in the "Define Constants" block below. Where double quotes are provided, insert
values between them. For most of the constants, the comments are exhaustive in explaining the needed contents.
Please see below detailed insructions on how to assign values to SEARCH_XPATH and CUSTOMVIEW_KEY.

Access your Orbis account and, in the left menu, select "Search". Here you can add all the needed search steps to identify
the companies for which you want to download data. When you are done, you can save this search clicking on the top-right
button. Now go to the "Load a search" menu, right-click on the name of your saved search, and select "Inspect Element".
Right-click the highlighted HTML code that appears and select the option to copy the XPath. You can then paste it as the
value of the SEARCH_XPATH constant below.

Now go back to your search and click on "View Results". In the page where results are loaded, it is possible to add and
remove variables from the final dataset to download. You can do this by clicking the "Add/remove columns" button on the 
right side of the page.
After choosing all the relevant variables, click "Apply". You will be back to the results page and a popup informs that 
you have successfully modified the view. On the top-right of the page, click on "Save as" and give a name to the new 
custom view.
Finally, simulate the download process by clicking on "Actions" and then on "Export companies". In the dialog window
that appears, right-click on the drop-down menu under "What would you like to export?" and select "Inspect Element".
Expand the highlighted HTML by clicking on the small arrow on the left. Many blocks starting with "<option" will
appear. Identify the one where "data-viewname" is equal to the name of your custom view. From that same block, you
will need to copy the string after "value=" and to paste it as the value of the CUSTOMVIEW_KEY constant below. You can 
do so by right-clicking on the block and selecting Copy Attribute. Then remember to delete the "value=" part after
pasting the string.

The rest of the constants for which you need to insert values below are more immediate and well described by the 
comments.

Note that in DOWNLOAD_DIR you have to indicate the path where the downloads will actually happen in your computer, as 
per your default settings. If you want to download files in a specific location, different from the default one, you
will have to set this manually in your Safari settings, as this code does not currently allow to remotely change it.


When all the constants are defined according to your needs, you can run the program without making any further change 
to the code below.

------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------


ADDITIONAL INFORMATION

The code running time strongly depends on connection speed. However, when a few hundred millions of observations need to 
be downloaded it can easily take several hours to execute the whole program. During this long running period, some errors 
which are not handled by the code might cause it to stop. It is your task to identify and to handle these rare exceptions 
in case they arise.

After the end of the program, from line 635, I have provided additional code snippets that might be useful to modify the
program according to your needs. These include, among others, further error handling possibilities and an option to send
yourself an email everytime a file is downloaded.

------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------------------------------------------------
'''




#-----------------------------------------------------------------------------------------------------------------------
 # DEFINE CONSTANTS

TOT_OBSERVATIONS = int  # Total number of firms for which you want to download data  (Please replace "int" with actual value)
STEP = int  # Number of observations you want to include in each .csv file  (Please replace "int" with actual value)

LOGIN_URL = ""  # URL of institution login page (NOT standard Orbis login page)
USER_PATH = ""  # Path where username is stored in your computer
PSW_PATH = ""  # Path where password is stored in your computer

SEARCH_XPATH = ''  # XPath of the button loading the custom search saved in the Orbis account

CUSTOMVIEW_KEY = ""  # Option value of the custom view choice in the export window

DOWNLOAD_DIR = ""  # Path where the .csv files are downloaded

TIMEOUT = int   # Number of seconds after which the download is retried (Please replace "int" with actual value)

#-----------------------------------------------------------------------------------------------------------------------














#-----------------------------------------------------------------------------------------------------------------------

 # IMPORT NECESSARY MODULES

import os
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import ElementNotInteractableException, TimeoutException

#-----------------------------------------------------------------------------------------------------------------------




#-----------------------------------------------------------------------------------------------------------------------

 # DEFINE RELEVANT FUNCTIONS


# Function to automatically login to Bocconi account
def auto_login(login_url):
     
    # Open the SSO login page
    driver.get(login_url)

    # Switch to the currently active window handle
    driver.switch_to.window(driver.current_window_handle)

    # Wait for the page to load
    time.sleep(2)

    # Find and double-click the "Students" button
    students_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="lachooser-container"]/div[2]/div/div[1]/div[1]/div/div[3]/span')))
    students_button.click()
    print("Student login started")

    # Wait for the page to load
    time.sleep(2)


    # Retrieve username
    with open(USER_PATH, "r") as user_file:
            username = user_file.read()

    # Fill in the username
    username_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "username")))
    username_field.send_keys(username)

    # Retrieve password
    with open(PSW_PATH, "r") as psw_file:
            password = psw_file.read()

    # Fill in the password
    password_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
    password_field.send_keys(password)


    # Submit the form
    submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/form/section/div[3]/button")))
    submit_button.click()
    print("Login successful")


    # Wait for the "Log in to Orbis" message to appear
    try:
        restart_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LoginForm"]/div/div/input')))
        # If the "Restart" button is found, click it
        restart_button.click()
        print("Open session with same account detected by Orbis. Session restarted.")

    except TimeoutException:
        # If the message doesn't appear within the timeout, continue with the next steps
        pass







# Function to load the custom search in Orbis
def load_search(search_xpath):

     # Find and click the "Load a Search" button
    loadsearch_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="tooltabload-search-section"]')))
    loadsearch_button.click()
    print('"Load a search" menu opened')

    # Wait for page loading
    time.sleep(5)

    # Find and click the "Italian Firms (All)" button
    itafirms_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_xpath)))
    itafirms_button.click()
    print("Search loaded: Italian Firms (All)")

    # Wait for page loading
    time.sleep(5)

    # Find and click the "View Results" button
    while True:
        
        try:
            results_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="search-summary"]/div/div/a')))
            results_button.click()
            
            print("Results loading...")
            break  # Exit the loop if the button is clicked successfully

        except ElementNotInteractableException:
            print("Results button is not interactable. Retrying...")
            
            # Refresh page
            driver.refresh()

            # Sleep for a short duration before retrying
            time.sleep(5)

    # Wait for page loading
    time.sleep(5)



# Function to download data in .csv files from the Orbis database
def download_data(start, end, index):

    # Find and click the "Actions" button
    actions_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/section[2]/div[1]/div[2]/div/div[2]/div[1]/ul/li[2]/a')))
    actions_button.click()
    print('"Actions" button found')

    # Find and click the "Export companies" button
    expcompanies_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '/html/body/section[2]/div[1]/div[2]/div/div[2]/div[1]/div/div/div[2]/a[2]/img')))
    expcompanies_button.click()
    print('"Export companies" button found')

    # Find the file type drop-down menu element
    filetype_menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="component_FormatTypeSelectedId"]')))
    # Select the CSV option by its value
    filetype_menu.send_keys("Custom.list.csv")

    # Find the view drop-down menu element
    view_menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="component_selectedFormatId"]')))
    # Select the "All Italian Firms" option by its value
    view_menu.send_keys(CUSTOMVIEW_KEY)

    # Find the companies drop-down menu element
    companies_menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="component_RangeOptionSelectedId"]')))
    # Create a Select object
    select = Select(companies_menu)
    # Select the "A range of companies" option by its value
    select.select_by_value("Range")

    # Find the input fields for the numbers
    input_field1 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="export-component-range"]/fieldset/input[1]')))
    input_field2 = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="export-component-range"]/fieldset/input[2]')))
    # Enter the numbers into the fields
    input_field1.send_keys(start)
    input_field2.send_keys(end)

    # Find the file name field and clear its current value after storing it
    text_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="component_FileName"]')))
    existing_text = text_field.get_attribute("value")
    text_field.clear()
    # Add a number at the beginning of the name
    new_filename = index + " " + existing_text
    text_field.send_keys(new_filename)

    # Find and click the "Export" button
    export_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="exportDialogForm"]/div[2]/a[2]')))
    export_button.click()

    return new_filename




# Function to verify that download is finished before continuing to next file
def end_download(directory_to_check, new_filename):

    # Wait for the download to complete
    download_directory = directory_to_check
    complete_filename = new_filename + ".csv"
    file_path = os.path.join(download_directory, complete_filename)

    # Wait for the download to complete
    WebDriverWait(driver, TIMEOUT).until(lambda x: os.path.exists(file_path))
        
    # Wait for page loading
    time.sleep(1)

    # Close the "Export" window
    close_button = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.XPATH, '/html/body/section[2]/div[6]/div[1]/img')))
    close_button.click()




# Function running a download loop starting from where the last download stopped
def loop_download(indices_to_download, search_xpath, download_dir, step):
    
    for i in indices_to_download:

        retry_count = 0  # Initialize retry count for each file
        max_retries = 5  # Maximum number of retries for each file

        while retry_count < max_retries:
            try:

                print(f"Downloading file n°{i}...")

                # Wait for the "Log in to Orbis" message to appear
                try:

                    # Wait for the "You currently have no active search" message to appear
                    try:
                        goto_search = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//a[contains(text(), "GO TO SEARCH")]')))
                        # If the "Go to search" button is found, click it
                        goto_search.click()
                        print('"No active search" error raised. Reloading search...')

                        # Load the custom search
                        load_search(search_xpath)

                    except TimeoutException:
                        # If the message doesn't appear within the timeout, continue with the next steps
                        pass


                    ok_button = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LoginForm"]/div/div/div/a')))
                    # If the "Ok" button is found, click it
                    ok_button.click()
                    print("Orbis unable to authenticate account. Reloading search...")

                    # Click the "Go to search" button
                    goto_search = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content"]/div/div/div/a')))
                    goto_search.click()
                    print("Search menu reached")

                    # Load the custom search
                    load_search(search_xpath)

                except TimeoutException:
                    # If the message doesn't appear within the timeout, continue with the next steps
                    pass


                # Get the current time before initiating the download
                start_time = time.time()

                # Download all data in .csv files with STEP lines each
                # Wait for the download to complete, then close the "Export" window
                end_download(download_dir, download_data(str(((i - 1) * step) + 1), str(i * step), str(i)))

                # Calculate the elapsed time for the download
                end_time = time.time()
                download_time = end_time - start_time
                print(f"File n°{i} downloaded in {download_time:.2f} seconds")

                # Exit the retry loop if the download succeeds
                break

            except TimeoutException:
                # If a TimeoutException occurs, retry the download
                print(f"Timeout occurred while downloading file {i}. Retrying...")
                
                # Refresh page
                driver.refresh()

                # Increment the retry count
                retry_count += 1

                # Sleep for a short duration before retrying
                time.sleep(5)

        else:
            # If maximum retries reached, move to the next file
            print(f"Unable to download file {i} after {max_retries} retries. Moving to the next file.")

            continue





# Function to extract the numeric index from the file name
def extract_index(filename):
    return int(filename.split()[0])




# Function to check which files are missing from the download directory
def check_files(lastfile, download_dir):

    # Define the range of indices you expect
    expected_indices = set(range(1, lastfile + 1))

    # List all files in the download directory
    all_files = os.listdir(download_dir)

    # Filter only CSV files and extract the indices
    indices = {extract_index(file) for file in all_files if file.endswith(".csv")}

    # Identify missing indices by finding the difference between expected and actual indices
    missing_indices = expected_indices - indices

    # Convert set of missing indices to a sorted list
    sorted_missing_indices = sorted(list(missing_indices))

    print("Missing files:", sorted_missing_indices)

    return sorted_missing_indices


#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------


 # COMPUTE NUMBER OF FILES NEEDED AND REMAINING OBSERVATIONS

penultimate_file = TOT_OBSERVATIONS // STEP
last_file = penultimate_file + 1

remaining_obs = TOT_OBSERVATIONS % STEP



#-----------------------------------------------------------------------------------------------------------------------

 # CHECK WHICH FILES STILL NEED TO BE DOWNLOADED

missing_files = check_files(last_file, DOWNLOAD_DIR)

#-----------------------------------------------------------------------------------------------------------------------




#-----------------------------------------------------------------------------------------------------------------------
 
while missing_files != []:

    print("Some files are missing. Proceding with the download...")


    # Path to WebDriver executable
    driver_path = "/System/Cryptexes/App/usr/bin/safaridriver"


    # Initialize the WebDriver
    driver = webdriver.Safari()
    #-----------------------------------------------------------------------------------------------------------------------





    #-----------------------------------------------------------------------------------------------------------------------

    # AUTO-LOGIN TO ORBIS

    auto_login(LOGIN_URL)

    #-----------------------------------------------------------------------------------------------------------------------



    #-----------------------------------------------------------------------------------------------------------------------
    # LOAD THE CUSTOM SEARCH

    load_search(SEARCH_XPATH)
    #-----------------------------------------------------------------------------------------------------------------------



    #-----------------------------------------------------------------------------------------------------------------------

    # LOOP TO DOWNLOAD DATA IN .csv FILES



    # Loop starting from where the last download stopped
    loop_download(missing_files, SEARCH_XPATH, DOWNLOAD_DIR, STEP)



    print("Main loop executed successfully. Proceding to last file...")


#-----------------------------------------------------------------------------------------------------------------------


    # LAST ITERATION (SINGLE) TO DOWNLOAD LAST .csv FILE


    # Execute only if not already downloaded
    if last_file in missing_files:

        last_startline = TOT_OBSERVATIONS - remaining_obs + 1  # Define starting line for last download

        retry_count = 0  # Initialize retry count
        max_retries = 5  # Maximum number of retries

        while retry_count < max_retries:
            try:
                # Get the current time before initiating the download
                start_time = time.time()

                # Download remaining data in a .csv file
                # Wait for the download to complete, then close the "Export" window
                end_download(DOWNLOAD_DIR, download_data(str(last_startline), str(TOT_OBSERVATIONS), str(last_file)))

                # Calculate the elapsed time for the download
                end_time = time.time()
                download_time = end_time - start_time
                print(f"File n°{last_file} downloaded in {download_time:.2f} seconds")

                # Exit the retry loop if the download succeeds
                break

            except TimeoutException:
                # If a TimeoutException occurs, retry the download
                print(f"Timeout occurred while downloading file {last_file}. Retrying...")
                
                # Refresh page
                driver.refresh()

                # Increment the retry count
                retry_count += 1

                # Sleep for a short duration before retrying
                time.sleep(5)

        else:
            # If maximum retries reached, move to the next file
            print(f"Unable to download file {last_file} after {max_retries} retries. Moving to final checks.")
    
    else:
        # If the last file has already been downloaded, continue with the next steps
        print("Last file already downloaded. Moving to final checks.")
        pass



    # Update list with missing files
    missing_files = check_files(last_file, DOWNLOAD_DIR)


    print("All downloads attempted without errors raised.")

#-----------------------------------------------------------------------------------------------------------------------

    # Close the browser
    driver.quit()
    print("Browser closed correctly.")

    # Wait for page loading
    time.sleep(5)

#-----------------------------------------------------------------------------------------------------------------------


else:

    print("All files have alredy been downloaded.")





print("Program finished.")


#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------
 # END OF PROGRAM






'''
# Server Error HTML:

<h1 data-debug-mode="false">Server Error</h1>
'''



'''
# Server Error handling:

from selenium.common.exceptions import NoSuchElementException
while True:
    try:
        # Check if the "Server Error" element is present
        error_message = driver.find_element_by_xpath('//h1[@data-debug-mode="false" and text()="Server Error"]')
        print("Encountered Server Error. Going back to the previous page...")
        
        driver.back()  # Navigate back to the previous page
        # Allow the page to reload completely
        time.sleep(5)

        driver.refresh()  # Refresh the page
        # Allow the page to reload completely
        time.sleep(5)

    except NoSuchElementException:
        # If the "Server Error" element is not found, break out of the loop
        break
'''




'''
# Option to send email after each download:

# Define your email as a constant
EMAIL = ""
# Find the email checkbox element and click it
checkbox = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="export-component-mailto"]/label')))
checkbox.click()
# Find the email field and clear its current value
email_field = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="component_EmailAddresses"]')))
email_field.clear()
# Insert email into the field
email = EMAIL
email_field.send_keys(email)
'''




'''
# Discover the path of the downloads directory:

import os
# Get the default download directory for the current user
download_directory = os.path.expanduser("~/Downloads")
# Print the path of the download directory
print("Download directory:", download_directory)
'''