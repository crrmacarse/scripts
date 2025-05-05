import os
import time
from slack_sdk import WebClient
from slack_sdk.errors import SlackApiError
from dotenv import load_dotenv
import argparse

load_dotenv()

def set_slack_status(status_text, status_emoji, status_expiration=0):
    try:
        slack_token = os.getenv("SLACK_OATH_TOKEN")
        if not slack_token:
            raise ValueError("API key not found in environment variables.")

        client = WebClient(token=slack_token)

        client.users_profile_set(
            profile={
            "status_text": status_text,
            "status_emoji": status_emoji,
            "status_expiration": int(time.time()) + status_expiration if status_expiration > 0 else 0
            }
        )
        print("Status updated successfully.")
    except SlackApiError as e:
        print(f"Error updating status: {e.response['error']}")


if __name__ == "__main__":
    try:
        parser = argparse.ArgumentParser(description="Slack Status Updater")

        parser.add_argument("--template", required=False, help="Template")

        args = parser.parse_args()
        template = args.template
        
        status_text = ""
        status_emoji = ""
        status_expiration = 0

        if template == "Lunch":
            status_text = "Lunch Break"
            status_emoji = ":knife_fork_plate:"
            status_expiration = 60 * 60
        elif template == "Break":
            status_text = "Quick Break"
            status_emoji = ":face_in_clouds:"
            status_expiration = 15 * 60
        else:
            status_text = input("Enter Status: ")
            status_emoji = input("Enter Emoji: ") or ""
            status_expiration = int(input("Enter Expiration (in seconds): "))
        
        set_slack_status(status_text, status_emoji, status_expiration)
    except Exception as e:
        print(f"Error{e}")