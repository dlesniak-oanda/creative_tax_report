import json
import os
from collections import OrderedDict
from datetime import datetime, timedelta

import requests
from dotenv import load_dotenv
from requests.auth import HTTPBasicAuth
from openpyxl import Workbook
from openpyxl.styles import Font

load_dotenv()

TASKS_COLUMNS = [
    'number',
    'Kod w systemie/Code in system',
    'Nazwa zadania/Task',
    'Krótki opis (pierwsze 200 znaków)/Short description (first 200 characters)',
    'TKP: Typ utworu (możemy przygotować katalog z którego pracownik wybierze jeden typ utworu, np. kod źródłowy, dokumentacja oprogramowania, grafika itd.) (in Polish)'
]


def get_env_variable(name, input_message):
    env_variable = os.environ.get(name, '')
    if not env_variable:
        env_variable = input(f'{input_message}\n')
    return env_variable


def get_jira_tasks(start_date, end_date):
    domain = "https://oandacorp.atlassian.net/"
    api_token = get_env_variable(name='JIRA_API_TOKEN', input_message="Provide Jira API Token")
    email = get_env_variable(name='EMAIL', input_message="What is your email addres connected to jira?")
    auth = HTTPBasicAuth(email, api_token)
    headers = {
        "Accept": "application/json"
    }
    query = {
        'jql': 'Assignee = currentUser() AND status = "Done"',
    }
    if start_date and end_date:
        # updatedDate  >=  "2018/10/01" and updatedDate   <= "2018/10/31"
        query['jql'] += f' AND updatedDate  >=  "{start_date}" and updatedDate   <= "{end_date}"'
    response = requests.request(
        "GET",
        f"{domain}/rest/api/3/search",
        headers=headers,
        params=query,
        auth=auth
    )
    return json.loads(response.text)


def get_reporting_period():
    date_now = datetime.now().date()
    default_reporting_start, default_reporting_end = get_start_end_month_day(month=date_now.month, year=date_now.year)
    while True:
        reporting_period = input(
            f'Okres raportowania/Reporting period format mm-yyyy (leave blank for date from {default_reporting_start} to {default_reporting_end})\n')
        if not reporting_period:
            start_date, end_date = default_reporting_start, default_reporting_end
            break
        start_date, end_date = date_from_input(reporting_period)
        if start_date and end_date:
            break
        print(f'Invalid date {reporting_period}, try again')
    return start_date, end_date


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
        year_future = year + 1
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


def get_header(data, start_date, end_date):
    employee_id = get_env_variable('EMPLOYEE_ID', 'Employee ID')
    job_position = get_env_variable('JOB_POSITION', 'Stanowisko/Job position')

    report_date = f'{start_date.day}-{end_date.day}.{start_date.month}.{start_date.year}'

    header = OrderedDict()
    header['Imie i nazwisko/Name and surname'] = data['issues'][0]['fields']['assignee']['displayName']
    header['Employee ID'] = employee_id
    header['Stanowisko/Job position'] = job_position,
    header['Okres raportowania/Reporting period'] = report_date
    header['Data zlozenia raportu/Submission date'] = datetime.now().date().strftime("%d.%m.%Y")
    return header


def generate_xlsx(header, tasks):
    workbook = Workbook()
    sheet = workbook.active
    populated_row_number = xlsx_populate_header(header, sheet)
    populated_row_number += 1
    column_letters = list("ABCDEFGHIJK")
    populate_body(column_letters, populated_row_number, sheet, tasks, workbook)
    return workbook


def populate_body(column_letters, populated_row_number, sheet, tasks, workbook):
    for column_name, column_letter in zip(TASKS_COLUMNS, column_letters):
        sheet[f'{column_letter}{populated_row_number}'] = column_name
        sheet[f'{column_letter}{populated_row_number}'].font = Font(bold=True)
    for task in tasks:
        populated_row_number += 1
        for column_letter, column_name in zip(column_letters, TASKS_COLUMNS):
            sheet[f'{column_letter}{populated_row_number}'] = task[column_name]
    adjust_columns_width(sheet)


def adjust_columns_width(ws):
    for column_cells in ws.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        ws.column_dimensions[column_cells[0].column_letter].width = length


def xlsx_populate_header(header, sheet, row=1):
    try:
        for key, value in header.items():
            sheet[f'A{row}'] = key
            sheet[f'B{row}'] = value
            row += 1
    except ValueError:
        sheet[f'B{row}'] = value[0]
        row += 1
    return row


start_date, end_date = get_reporting_period()
data = get_jira_tasks(start_date, end_date)
if data.get('issues'):
    tasks = get_tasks_rows(data)
    header = get_header(data, start_date, end_date)
    workbook = generate_xlsx(header=header, tasks=tasks)
    workbook.save(filename=f"creative-tax-{start_date}.xlsx")

    for task in tasks:
        print(task)
    print(f"{len(tasks)} tasks,")
    print([task['Kod w systemie/Code in system'] for task in tasks])
else:
    print(data.get('errorMessages', 'Unexpected error, no issues'))
