import array
from dataclasses import dataclass
from dotenv import load_dotenv
import os
import re
import requests
import xlsxwriter


@dataclass
class ReportEntry:
    """
    ReportEntry dataclass stores the math placement result for one student.
    """
    student_name: str = ''
    student_sis_id: str = ''
    attempts: int = 0
    best_attempt: int = 0
    subscore_q01_to_q08: int = 0
    subscore_q09_to_q16: int = 0
    subscore_q17_to_q24: int = 0
    subscore_q25_to_q32: int = 0
    subscore_q33_to_q36: int = 0
    total_score: int = 0
    placement: str = ''


load_dotenv()
canvas_access_token = os.getenv('CANVAS_ACCESS_TOKEN')


def get_quiz_url():
    """
    Prompts the user to input the URL of a Canvas quiz.
    """
    print('Dominican University')
    print('Mathematics Placement Exam Reporting Tool')
    print()
    print('Enter the Canvas URL of the Mathematics Placement Exam.')
    return input('> ')


def parse_quiz_url(quiz_url):
    """
    Parses the quiz URL to get the course ID, quiz ID, and assignment ID.
    """
    quiz_url_format = '\Ahttps://dominicanu.instructure.com/courses/[1-9][0-9]*/quizzes/[1-9][0-9]*\Z'
    if re.findall(quiz_url_format, quiz_url):
        course_id = re.search('/courses/(.+?)/quizzes/', quiz_url).group(1)
        print(f'Parsed course ID is >{course_id}<.')
        quiz_id = re.search('/quizzes/(.*)$', quiz_url).group(1)
        print(f'Parsed quiz ID is >{quiz_id}<.')
        url = 'https://dominicanu.instructure.com/api/v1/courses/' + course_id + '/quizzes/' + quiz_id
        payload = {}
        headers= {'Authorization': 'Bearer ' + canvas_access_token}
        response = requests.request("GET", url, headers=headers, data=payload)
        print(f'Response status code is >{response.status_code}<.')
        if (response.status_code == 200):
             return course_id, quiz_id, str(response.json()['assignment_id'])
            
    print('The format of the URL provided is not valid.')
    print('I expected something like "https://dominicanu.instructure.com/courses/[number]/quizzes/[number]".')
    return '0', '0', '0'


def initialize_workbook():
    """
    Creates an Excel workbook, a worksheet, and sets up column headers for the report.
    """
    workbook = xlsxwriter.Workbook('Math Placements ' + quiz_id + '.xlsx')
    worksheet = workbook.add_worksheet('Math Placements ' + quiz_id)
    worksheet.write(0, 0, 'Student Name')
    worksheet.write(0, 1, 'Student ID')
    worksheet.write(0, 2, 'Attempts')
    worksheet.write(0, 3, 'Best Attempt')
    worksheet.write(0, 4, 'Q1–Q8 Subscore')
    worksheet.write(0, 5, 'Q9–Q16 Subscore')
    worksheet.write(0, 6, 'Q17–Q24 Subscore')
    worksheet.write(0, 7, 'Q25–Q32 Subscore')
    worksheet.write(0, 8, 'Q33–Q36 Subscore')
    worksheet.write(0, 9, 'Total Score')
    worksheet.write(0, 10, 'Placement')
    return workbook, worksheet, 0, 1


def get_submissions():
    """
    Creates a list of student submissions for a given assignment.
    """
    submissions = []

    # Get the first page of submissions
    url = 'https://dominicanu.instructure.com/api/v1/courses/' + course_id + '/assignments/' + assignment_id + '/submissions?include[]=submission_history'
    payload = {}
    headers= {'Authorization': 'Bearer ' + canvas_access_token}
    response = requests.request("GET", url, headers=headers, data=payload)
    if (response.status_code == 200):
        raw_responses = response.json()
        for raw_response in raw_responses:
            submissions.append(raw_response)

    # If there are additional pages of submissions, get those as well
    while 'next' in response.links:
        url = response.links['next']['url']
        payload = {}
        headers= {'Authorization': 'Bearer ' + canvas_access_token}
        response = requests.request("GET", url, headers=headers, data=payload)
        if (response.status_code == 200):
            raw_responses = response.json()
            for raw_response in raw_responses:
                submissions.append(raw_response)

    return submissions


def process_submission(submission):
    report_entry = ReportEntry()

    # Get the student's name and ID number
    url = 'https://dominicanu.instructure.com/api/v1/users/' + str(submission['user_id'])
    payload = {}
    headers= {'Authorization': 'Bearer ' + canvas_access_token}
    user = requests.request("GET", url, headers=headers, data=payload).json()
    report_entry.student_name = str(user['sortable_name'])
    report_entry.student_sis_id = str(user['sis_user_id'])

    # Record the number of attempts made by the student 
    report_entry.attempts = submission['attempt']

    # Tally the placement exam subscores
    for quiz_attempt in submission['submission_history']:
        subscores = array.array('I', [0, 0, 0, 0, 0])
        total_score = 0

        for idx, question in enumerate(quiz_attempt['submission_data']):
            if question['correct']:
                total_score += 1

                if 0 <= idx <= 7:
                    subscores[0] += 1
                elif 8 <= idx <= 15:
                    subscores[1] += 1
                elif 16 <= idx <= 23:
                    subscores[2] += 1
                elif 24 <= idx <= 31:
                    subscores[3] += 1
                else:
                    subscores[4] += 1

        if (subscores[4] >= 2) and (total_score >= 29):
            placement = 'MATH 261'
        elif ((subscores[0] >= 6) and (subscores[1] >= 6) and (subscores[2] >=6)) or (total_score >= 23):
             placement = 'MATH 250'
        elif ((subscores[0] >= 6) and (subscores[1] >= 6)) or (total_score >= 16):
             placement = 'MATH 130/150/170'
        elif (subscores[0] >=6) or (total_score >= 12):
            placement = 'MATH 120'
        else:
            placement = 'MATH 090'

        if (placement > report_entry.placement) or ((placement == report_entry.placement) and (total_score > report_entry.total_score)):
            report_entry.subscore_q01_to_q08 = subscores[0]
            report_entry.subscore_q09_to_q16 = subscores[1]
            report_entry.subscore_q17_to_q24 = subscores[2]
            report_entry.subscore_q25_to_q32 = subscores[3]
            report_entry.subscore_q33_to_q36 = subscores[4]
            report_entry.total_score = total_score
            report_entry.placement = placement
            report_entry.best_attempt = quiz_attempt['attempt']

    return report_entry


quiz_url = get_quiz_url()
course_id, quiz_id, assignment_id = parse_quiz_url(quiz_url)

if course_id != '0' and assignment_id != '0':
    workbook, worksheet, col, row = initialize_workbook()
    submissions = get_submissions()

    for submission in submissions:
        if submission['workflow_state'] == 'graded':
            report_entry = process_submission(submission)

            print('Student Name: ' + report_entry.student_name)
            print('Student ID: ' + report_entry.student_sis_id)
            print('Attempts: ' + str(report_entry.attempts))
            print('Best Attempt: ' + str(report_entry.best_attempt))
            print('Q01 - Q08 Subscore: ' + str(report_entry.subscore_q01_to_q08))
            print('Q09 - Q16 Subscore: ' + str(report_entry.subscore_q09_to_q16))
            print('Q17 - Q24 Subscore: ' + str(report_entry.subscore_q17_to_q24))
            print('Q25 - Q32 Subscore: ' + str(report_entry.subscore_q25_to_q32))
            print('Q33 - Q36 Subscore: ' + str(report_entry.subscore_q33_to_q36))
            print('Total Score: ' + str(report_entry.total_score))
            print('Placement: ' + report_entry.placement)

            worksheet.write(row, col, report_entry.student_name)
            col += 1
            worksheet.write(row, col, int(report_entry.student_sis_id))
            col += 1
            worksheet.write(row, col, report_entry.attempts)
            col += 1
            worksheet.write(row, col, report_entry.best_attempt)
            col += 1
            worksheet.write(row, col, report_entry.subscore_q01_to_q08)
            col += 1
            worksheet.write(row, col, report_entry.subscore_q09_to_q16)
            col += 1
            worksheet.write(row, col, report_entry.subscore_q17_to_q24)
            col += 1
            worksheet.write(row, col, report_entry.subscore_q25_to_q32)
            col += 1
            worksheet.write(row, col, report_entry.subscore_q33_to_q36)
            col += 1
            worksheet.write(row, col, report_entry.total_score)
            col += 1
            worksheet.write(row, col, report_entry.placement)
            col = 0
            row += 1

    workbook.close()