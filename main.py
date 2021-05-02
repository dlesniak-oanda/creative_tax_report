import json
import os
from collections import OrderedDict
from datetime import datetime, timedelta

import requests
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook

TASKS_COLUMNS = [
    'number',
    'Kod w systemie/Code in system',
    'Nazwa zadania/Task',
    'Krótki opis (pierwsze 200 znaków)/Short description (first 200 characters)',
    'TKP: Typ utworu (możemy przygotować katalog z którego pracownik wybierze jeden typ utworu, np. kod źródłowy, dokumentacja oprogramowania, grafika itd.) (in Polish)'
]


def get_jira_tasks():
    domain = "https://oandacorp.atlassian.net/"
    api_token = os.environ.get('JIRA_API_TOKEN', input("Provide Jira API Token"))
    email = os.environ.get('EMAIL', input("What is your email addres"))
    auth = HTTPBasicAuth(email, api_token)
    headers = {
        "Accept": "application/json"
    }
    query = {
        'jql': 'Assignee = currentUser() AND status = "Done"',
    }
    response = requests.request(
        "GET",
        f"{domain}/rest/api/3/search",
        headers=headers,
        params=query,
        auth=auth
    )
    return json.loads(response.text)


def get_short_description(task):
    short_description = ''
    try:
        for content_block in task['fields']['description']['content']:
            for content in content_block['content']:
                if content_block['type'] == 'paragraph' and content['type'] == 'text':
                    short_description += content['text']
                    if len(short_description) > 200:
                        break
            if len(short_description) > 200:
                break
    except:
        return ''
    return short_description[:200]


def get_tasks_rows(data):
    tasks = []
    for number, task in enumerate(data['issues']):
        tasks.append({
            'number': number + 1,
            'Kod w systemie/Code in system': task['key'],
            'Nazwa zadania/Task': task['fields']['summary'],
            # ['fields']['status']['name'] == In Progress
            'Krótki opis (pierwsze 200 znaków)/Short description (first 200 characters)': get_short_description(task),
            'TKP: Typ utworu (możemy przygotować katalog z którego pracownik wybierze jeden typ utworu, np. kod źródłowy, dokumentacja oprogramowania, grafika itd.) (in Polish)': 'kod źródłowy',
        })
    return tasks


def get_start_end_month_day(month, year):
    month, year = int(month), int(year)
    start_date = datetime.now().replace(day=1, month=month, year=year).date()
    if month == 12:
        month_future = 1
        year_future = +1
    else:
        month_future = month + 1
        year_future = year
    end_date = (start_date.replace(day=1, month=month_future, year=year_future) - timedelta(days=1))
    return start_date, end_date


def date_from_input(raw_date):
    try:
        month, year = raw_date.split('-')
        return get_start_end_month_day(month=month, year=year)
    except:
        return None, None


def get_header(data):
    date_now = datetime.now().date()
    default_reporting_start, default_reporting_end = get_start_end_month_day(month=date_now.month, year=date_now.year)
    employee_id = input('Employee ID\n')
    job_position = input('Stanowisko/Job position\n')
    while True:
        reporting_period = input(
            f'Okres raportowania/Reporting period format mm-yyyy (leave blank for {default_reporting_start}-{default_reporting_end})\n')
        if not reporting_period:
            start_date, end_date = default_reporting_start, default_reporting_end
            break
        start_date, end_date = date_from_input(reporting_period)
        if start_date and end_date:
            break
        print(f'Invalid date {reporting_period}, try again')

    submission_date = input(f'Data zlozenia raportu/Submission date leave blank for {date_now}\n')
    report_date = f'{start_date.day}-{end_date.day}.{start_date.month}.{start_date.year}'
    submission_date = submission_date.strftime("%d.%m.%Y") if submission_date else date_now.strftime("%d.%m.%Y")

    header = OrderedDict()
    header['Imie i nazwisko/Name and surname'] = data['issues'][0]['fields']['assignee']['displayName']
    header['Employee ID'] = employee_id
    header['Stanowisko/Job position'] = job_position,
    header['Okres raportowania/Reporting period'] = report_date
    header['Data zlozenia raportu/Submission date'] = submission_date
    return header


def generate_xlsx(header, tasks):
    workbook = Workbook()
    sheet = workbook.active
    populated_row_number = xlsx_populate_header(header, sheet)

    column_letters = list("ABCDEFGHIJK")
    for column_name, column_letter in zip(TASKS_COLUMNS, column_letters):
        sheet[f'{populated_row_number}{column_letter}'] = column_name
    populated_row_number += 1
    for task in tasks:
        for column in TASKS_COLUMNS:
            populated_row_number

    workbook.save(filename=f"creative-tax-{datetime.now().date()}.xlsx")


def xlsx_populate_header(header, sheet):
    row = 0
    for key, value in header:
        sheet[f'A{row}'] = key
        sheet[f'B{row}'] = value
        row += 1
    return row


data = get_jira_tasks()
if data.get('issues'):
    tasks = get_tasks_rows(data)
    header = get_header(data)
    generate_xlsx(header=header, tasks=tasks)
    for task in tasks:
        print(task)
    print(f"{len(tasks)} tasks,")
    print([task['Kod w systemie/Code in system'] for task in tasks])
else:
    print(data.get('errorMessages', 'Unexpected error, no issues'))
