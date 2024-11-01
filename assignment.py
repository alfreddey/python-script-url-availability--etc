import requests
import subprocess
import os
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
import win32com.client
from dotenv import load_dotenv

load_dotenv()

# Set up logging to record what happens in the script
logging.basicConfig(filename='system_health.log', format='%(asctime)s %(levelname)-8s %(message)s', level=logging.INFO)

# Email settings
SMTP_SERVER = os.getenv('SMTP_SERVER')
SMTP_PORT = os.getenv('SMTP_PORT')
SMTP_USER = os.getenv('SMTP_USER')
SMTP_PASSWORD = os.getenv('SMTP_PASSWORD')
RECIPIENTS = os.getenv('RECIPIENTS')

# URL and server info
URL_TO_CHECK = os.getenv('URL_TO_CHECK')
SERVER_IP = os.getenv('SERVER_IP')
SCHEDULED_TASK_NAME = os.getenv('SCHEDULED_TASK_NAME')

# Function to send emails when something happens
def send_email(subject, body):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = ', '.join(RECIPIENTS)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, RECIPIENTS, msg.as_string())
        logging.info(f"Sent email: {subject}")
    except Exception as e:
        logging.error(f"Couldn't send email: {e}")

# Function to check if a URL is reachable
def check_url():
    try:
        response = requests.get(URL_TO_CHECK, timeout=10)
        if response.status_code == 200:
            subject = "URL is Working!"
            body = f"The URL `{URL_TO_CHECK}` was reachable at {datetime.now()}."
            send_email(subject, body)
            logging.info(f"URL check successful: {URL_TO_CHECK}")
        else:
            raise Exception(f"Got status code {response.status_code}")
    except Exception as e:
        subject = "URL Problem"
        body = f"The URL `{URL_TO_CHECK}` could not be reached at {datetime.now()}. Error: {e}"
        send_email(subject, body)
        logging.error(f"URL check failed: {e}")

# Function to ping a server to check if it’s online
def ping_server():
    try:
        result = subprocess.run(['ping', '-n', '1', SERVER_IP], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode == 0:
            subject = "Server is Online!"
            body = f"The server at IP `{SERVER_IP}` responded to ping at {datetime.now()}."
            send_email(subject, body)
            logging.info(f"Ping successful: {SERVER_IP}")
        else:
            raise Exception(f"Ping failed with code {result.returncode}")
    except Exception as e:
        subject = "Server Unreachable"
        body = f"The server at IP `{SERVER_IP}` couldn't be reached at {datetime.now()}. Error: {e}"
        send_email(subject, body)
        logging.error(f"Ping failed: {e}")

# Function to check if a Windows Scheduled Task has run successfully
def check_scheduled_task():
    try:
        scheduler = win32com.client.Dispatch("Schedule.Service")
        scheduler.Connect()
        rootFolder = scheduler.GetFolder("\\")
        task = rootFolder.GetTask(SCHEDULED_TASK_NAME)
        
        last_run_time = task.LastRunTime
        last_task_result = task.LastTaskResult

        if last_task_result == 0:
            subject = "Task Ran Successfully!"
            body = f"The task `{SCHEDULED_TASK_NAME}` ran at {last_run_time}."
        else:
            subject = "Task Failed"
            body = f"The task `{SCHEDULED_TASK_NAME}` did not run successfully. Last run time: {last_run_time}, Result: {last_task_result}."
        
        send_email(subject, body)
        logging.info(f"Checked task: {SCHEDULED_TASK_NAME}")
    except Exception as e:
        subject = "Task Check Error"
        body = f"Couldn't check the task `{SCHEDULED_TASK_NAME}`. Error: {e}"
        send_email(subject, body)
        logging.error(f"Task check failed: {e}")

# Main function to run all the checks
def main():
    check_url()
    ping_server()
    check_scheduled_task()

# Run the checks if this script is run directly
if __name__ == "__main__":
    main()
