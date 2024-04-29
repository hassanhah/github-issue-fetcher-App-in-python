import os
import requests
from openpyxl import Workbook

def fetch_issues(repo_owner, repo_name):
    url = f"https://api.github.com/repos/{repo_owner}/{repo_name}/issues"
    try:
        response = requests.get(url)
        response.raise_for_status()  # Raise exception for 4xx or 5xx status codes
        issues = response.json()
        issue_list = []
        for issue in issues:
            issue_number = issue['number']
            issue_title = issue['title']
            issue_author = issue['user']['login']
            issue_list.append((issue_number, issue_title, issue_author))
        return issue_list
    except requests.exceptions.RequestException as e:
        print(f"Failed to fetch issues from GitHub API: {e}")
        return None

def write_to_excel(issue_list, folder, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(['Issue Number', 'Title', 'Author'])
    for issue in issue_list:
        ws.append(issue)
    try:
        if not os.path.exists(folder):
            os.makedirs(folder)
        file_path = os.path.join(folder, filename)
        wb.save(file_path)
        print(f"Issues written to {file_path}.")
    except PermissionError:
        print(f"Permission denied: Unable to write to {file_path}.")
    except Exception as e:
        print(f"An error occurred while writing to {file_path}: {e}")

if __name__ == "__main__":
    repo_owner = input("Enter the repository owner's name: ")
    repo_name = input("Enter the repository name: ")
    folder_name = input("Enter the folder name to store the output file: ")
    
    issues_list = fetch_issues(repo_owner, repo_name)
    if issues_list is not None:
        filename = f"{repo_owner}_{repo_name}_issues.xlsx"
        write_to_excel(issues_list, folder_name, filename)
