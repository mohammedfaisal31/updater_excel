import requests
import pandas as pd
import schedule
import time
import psutil  # For checking if Excel files are open

def is_file_open(file_path):
    for process in psutil.process_iter():
        try:
            if file_path.lower() in process.cmdline() or file_path.lower() in " ".join(process.cmdline()):
                return True
        except (psutil.AccessDenied, psutil.NoSuchProcess, psutil.ZombieProcess):
            pass
    return False

def update_excel_files():
    # Specify the file paths for Excel files
    voted_excel_file = 'voters_who_voted.xlsx'
    not_voted_excel_file = 'voters_who_did_not_vote.xlsx'

    # Check if the Excel files are open
    if is_file_open(voted_excel_file) or is_file_open(not_voted_excel_file):
        print("Excel files are open. Unable to update.")
    else:
        # Send a GET request to the API URL
        api_url = "https://lyxnlabsapi.online/api/getVotersData"
        response = requests.get(api_url)

        # Check if the request was successful (status code 200)
        if response.status_code == 200:
            # Parse the JSON response into DataFrames
            data = response.json()
            df = pd.DataFrame(data)

            # Separate the data into voters who have voted and those who have not voted
            voted = df.dropna(subset=['candidate_first_name'])
            not_voted = df[df['candidate_first_name'].isna()]

            # Export the data to Excel files
            voted.to_excel(voted_excel_file, index=False)
            not_voted.to_excel(not_voted_excel_file, index=False)

            print("Data exported to Excel files.")
        else:
            print("Failed to retrieve data from the API.")

# Run the update function immediately
update_excel_files()


