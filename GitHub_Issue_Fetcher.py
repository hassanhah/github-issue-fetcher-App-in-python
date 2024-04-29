import requests
from openpyxl import Workbook
import os

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

def write_to_excel(issue_list, filename):
    wb = Workbook()
    ws = wb.active
    ws.append(['Issue Number', 'Title', 'Author'])
    for issue in issue_list:
        ws.append(issue)
    try:
        wb.save(filename)
        print(f"Issues written to {filename}.")
    except PermissionError:
        print(f"Permission denied: Unable to write to {filename}.")
    except Exception as e:
        print(f"An error occurred while writing to {filename}: {e}")

if __name__ == "__main__":
    repo_owner = input("Enter the repository owner's name: ")
    repo_name = input("Enter the repository name: ")
    issues_list = fetch_issues(repo_owner, repo_name)
    if issues_list is not None:
        filename = f"{repo_owner}_{repo_name}_issues.xlsx"
        write_to_excel(issues_list, filename)
