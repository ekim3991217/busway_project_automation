from github import Github
import os
from dotenv import load_dotenv

# Load .env file
load_dotenv()

# Access your secret
token = os.getenv("GITHUB_TOKEN")
g = Github(token)

user = g.get_user()
repo_name = "busway_project_automation"
repo = user.create_repo(repo_name)
print(f"Created repository {repo.full_name}")
