import requests
import pandas as pd
import io
import base64
import os

def main(req: func.HttpRequest) -> func.HttpResponse:
    # GitHub repository details
    owner = 'cassadn'
    repo_name = 'ct_simple_tracker'
    file_path = '/Users/a1609186/Library/CloudStorage/OneDrive-BostonScientific/Clinical and Wonder Data Analysis/Clinical Engineering/CT Special Ops/ct_ops_tracker/ct_ops_workup_tracker.xlsx'
    file_name = os.path.basename(file_path)

    # Your GitHub PAT for authentication
    github_token = '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'

    # GitHub API endpoint for fetching file contents and updating file
    base_url = f'https://api.github.com/repos/{owner}/{repo_name}/contents'

    # Request headers with GitHub PAT for authentication
    headers = {
        'Authorization': f'token {github_token}',
        'Content-Type': 'application/json'
    }
        
    # Fetch the file content from GitHub
    response = requests.get(base_url, headers=headers)

    if response.status_code == 200:
        # Find the 'ct_ops_workup_tracker.xlsx' file in the list of files
        file_info = next(file for file in response.json() if file['name'] == file_name)

        # make the sha 
        sha = file_info['sha']
        download_url = file_info['download_url']
        
        # Read the local file
        file_dir = os.path.dirname(file_path)
        os.chdir(file_dir)
        local_content = pd.read_excel(file_name)

        # Convert local to base64 to pass into GitHub
        csv_content = local_content.to_csv(index=False)
        encoded_content = base64.b64encode(csv_content.encode()).decode()

        # Fetch the file content of 'ct_ops_workup_tracker.xlsx' from GitHub
        file_response = requests.get(download_url, headers=headers)

        if file_response.status_code == 200:
            github_content = file_response.content
            content = pd.read_excel(io.BytesIO(github_content))

            # Check if local file content differs from GitHub file content
            if not local_content.equals(content):
                # Data for the PUT request to update the file
                data = {
                    'message': 'Update Excel file',
                    'content': encoded_content,
                    'sha': sha
                }
                
                
                # Send the PUT request to update the file on GitHub
                update_response = requests.put(file_info['url'], headers=headers, json=data)

                if update_response.status_code == 200:
                    print('File updated successfully on GitHub!')
                    
                else:
                    print('Error updating file:', update_response.text)
                    
            else:
                print('No changes detected in the local file.')
        else:
            print('Error fetching file content from GitHub:', file_response.text)
    else:
        print('Error fetching file details from GitHub:', response.text)
