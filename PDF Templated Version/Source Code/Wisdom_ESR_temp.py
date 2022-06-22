import os
import sys
from typing import Tuple
from numpy import int64
from scipy.stats import percentileofscore
import pandas as pd
from pandas.core.series import Series
from reportlab.platypus import Paragraph, Image, Table
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet
from utils.utils import PDFItem, PDFPage
from utils.static import PAGE_WIDTH, PAGE_HEIGHT, MARGIN_LEFT, MARGIN_TOP
from reportlab.platypus import Paragraph, Image, Table
from reportlab.lib.units import cm
# from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from typing import Tuple
from utils.utils import PDFItem, PDFPage
from utils.data_operations import get_percent_of_attempted_questions, get_hist
from utils.utils import PDF
from time import perf_counter, sleep
import typer

PAGE_HEIGHT = 450
PAGE_WIDTH = 800

def handleError(msg):
    typer.secho(msg, fg=typer.colors.BRIGHT_RED)
    sleep(3)
    sys.exit()


def get_data(filename: str, sheet: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
    head = ['Student No', 'Round', 'First Name of Candidate',
            'Last Name of Candidate', 'Full Name of Candidate',
            'Registration', 'Grade', 'Name of school', 'Gender',
            'Date of Birth', 'City of Residence', 'Date and time of test',
            'Country of Residence', 'Question No.', 'What you marked',
            'Correct Answer', 'Outcome (Correct/Incorrect/Not Attempted)',
            'Score if correct', 'Your score']
    converters = {
        'Candidate No. (Need not appear on the scorecard)': int, 
        'Your score': int, 'Date and time of test':str, 
        "Date of Birth":str, 'Registration': pd.Int64Dtype
        }
    try:
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            data = pd.read_excel(filename, sheet_name=sheet, converters=converters)
            data = data.iloc[:, [i for i in range(len(head))]]
        elif filename.endswith('.csv'):
            data = pd.read_csv(filename, sheet_name=sheet, converters=converters)
        else:
            raise TypeError('Invalid file extension.')
    except TypeError:
        handleError("Invalid File extension!")
    except:
        handleError("Some unexpected error with file!")
    data.columns = head
    data = data.dropna(axis=1, how="all")
    data = data.dropna(axis=0, thresh=6)
    data = data.fillna(value=" ")
    data['Question No.'] = data['Question No.'].str.strip()
    ROUND = int(data.iloc[10]['Round'])
    data.drop('Round', inplace=True, axis=1)
    if data.iloc[0]['Your score'] == 'Your score':
        data = data.iloc[1:]
    data.insert(1, 'Attempt Status', 'Attempted')
    for index, row in data.iterrows():
        if row['Outcome (Correct/Incorrect/Not Attempted)'] in ['Correct', 'Incorrect']:
            data['Attempt Status'][index] = 'Attempted'
        else:
            data['Attempt Status'][index] = 'Unattempted'
    clmns = ['Student No', 'First Name of Candidate',
             'Last Name of Candidate', 'Full Name of Candidate', 'Registration',
             'Grade', 'Name of school', 'Gender', 'Date of Birth',
             'Date and time of test', 'City of Residence', 'Country of Residence',
             'Question No.', 'What you marked', 'Correct Answer',
             'Outcome (Correct/Incorrect/Not Attempted)', 'Score if correct',
             'Your score']
    data['Date of Birth'] = pd.to_datetime(data['Date of Birth'])
    data['Date and time of test'] = pd.to_datetime(data['Date and time of test'], format="%b %d-6 %Y")
    static = data[['Student No', 'First Name of Candidate',
                   'Last Name of Candidate', 'Full Name of Candidate', 'Registration',
                   'Grade', 'Name of school', 'Gender', 'Date of Birth',
                   'Date and time of test', 'City of Residence', 'Country of Residence', ]].drop_duplicates(subset=['Student No', 'Full Name of Candidate'])
    const = data[['Question No.', 'Attempt Status', 'What you marked', 'Correct Answer',
                 'Outcome (Correct/Incorrect/Not Attempted)', 'Score if correct',
                 'Your score'] + ['Country of Residence', 'Student No']]
    static = static.astype({'Student No': int, 'Registration': int64, 'Grade': int})
    const = const.astype({'Student No': int, 'Your score': int, 'Score if correct': int})
    return static, const, ROUND

typer.secho("Press enter for default values!", fg=typer.colors.BRIGHT_YELLOW)
FILE = input("Enter the name of data file: (defaults to ./Dummy Data for final assignment.xlsx)\n")
if FILE == '':
    FILE = './Dummy Data for final assignment.xlsx'
try:    
    sheet_names = pd.ExcelFile(FILE).sheet_names
except:
    handleError("Invalid File location!")
typer.secho('Available Sheets '+ str(sheet_names), fg=typer.colors.BRIGHT_YELLOW)
SHEET = input("Enter the name of sheet in the file : (defaults to Sheet1 )\n")
if SHEET == '':
    SHEET = 'Sheet1'
else:
    if SHEET not in sheet_names:
        handleError("Sheet does not exist!")
PICS = input("Enter the location of student images : (defaults to ./Pics for assignment)\n")
if PICS == '':
    PICS = './Pics for assignment'


static, const, ROUND = get_data(FILE, sheet=SHEET)

typer.secho("Round detected is "+ str(ROUND), fg=typer.colors.BRIGHT_YELLOW)
user_round = input("Enter the round number if need to change (i.e. 1, 2 or 3)\n")
if user_round != '':
    try:
        ROUND = int(user_round)
    except:
        handleError("Value for Round is not correct!")
ques_group = const.groupby('Question No.')

stu_group = const.groupby('Student No')
Q_NO = len(ques_group)

PAGE_HEIGHT += int(19.5 * Q_NO)

t = perf_counter()
PAGE_TOP = PAGE_HEIGHT - 2 * MARGIN_TOP
PAGE_TOP_2 = PAGE_TOP + 100
HEADING = getSampleStyleSheet()['Heading3']

PAGE_TOP = PAGE_HEIGHT - 2 * MARGIN_TOP

FINAL_PERCENTILE = 0

heading = getSampleStyleSheet()['Heading3']


def get_page_1(stu_details, stu_record, img_loc):
    Page1 = PDFPage()
    background = Image(f"{REQUIREDFILES}/back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT, hAlign='CENTER')
    title = Paragraph("<font size=12><b>INTERNATIONAL MATHS OLYMPIAD CHALLENGE</b></font>")
    logo = Image(f"{REQUIREDFILES}/logo.png", width=7 * cm + 20, height=3.4 * cm,
                 hAlign='LEFT')
    headingl = Paragraph(
        text=f"<b>Round I - Enhanced Score Report</b>:  {stu_details['Full Name of Candidate'].to_string(index=False)}<br/>\nReg Number: {stu_details['Registration'].to_string(index=False)}")
    stu_pic = Image(f"{img_loc}", width=5 * cm, height=4 * cm,
                    hAlign="RIGHT")
    student_name = Paragraph(
        text=f"<b><font size=12>Round {ROUND} performance of {stu_details['Full Name of Candidate'].to_string(index=False)}</font></b>", )
    dob = stu_details['Date of Birth'].dt.strftime("%d %b %Y")
    dot = stu_details['Date and time of test'].dt.strftime('%d %b %Y')

    student_detail = (
        ("Grade", f"{str(int(stu_details['Grade']))}", "", "Registration No. ", f"{stu_details['Registration'].to_string(index=False)}"),
        ("School Name", f"{stu_details['Name of school'].to_string(index=False)}", "", "Gender", f"{stu_details['Gender'].to_string(index=False)}"),
        ("City Of Residence", f"{stu_details['City of Residence'].to_string(index=False)}", "", "Date of Birth",
         f"{dob.to_string(index=False)}"),
        ("Country Of Residence", f"{stu_details['Country of Residence'].to_string(index=False)}", "", "Date Of Test",
         f"{dot.to_string(index=False)}")
    )
    tbl_style = (
        ('FONT', (0, 0), (0, -1), "Helvetica-Bold"),
        ('FONT', (3, 0), (3, -1), "Helvetica-Bold"),
        ('INNERGRID', (0, 0), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 0), (-1, -1), 1, (0, 0, 0)),
        ('BOX', (0, 0), (1, -1), 1, (0, 0, 0)),
        ('BOX', (-2, 0), (-1, -1), 1, (0, 0, 0)),
    )
    table = Table(student_detail, style=tbl_style, colWidths=(None, 5 * cm, 1 * cm, None, 5 * cm))
    stf = Paragraph(text="Section 1", style=heading)
    desc = Paragraph(
        text=f"This section describes {stu_details['First Name of Candidate'].to_string(index=False)}'s performance v/s the Test in Grade {str(int(stu_details['Grade']))}")
    report_table_data = [
        ("Question No.", 'Attempt \nStatus', f"  {stu_details['First Name of Candidate'].to_string(index=False)}'s  \n  Choice  ", 'Correct\nAnswer',
         '  Outcome  ', 'Score if\ncorrect', f"  {stu_details['First Name of Candidate'].to_string(index=False)}'s  \n  Score  "),
    ]

    for _, row in stu_record.iterrows():
        report_table_data.append((
            row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
            row['Outcome (Correct/Incorrect/Not Attempted)'],
            row['Score if correct'], row['Your score']
        ))
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table1 = Table(report_table_data, style=style)
    tscore = Paragraph(text=f"<b><i>Total Score</i>: {stu_record['Your score'].sum()}</b>")
    data = (
        PDFItem(background, 0, 0),
        PDFItem(title, 245, PAGE_TOP_2 - 1 * cm - 100),
        PDFItem(headingl, MARGIN_LEFT + 6, PAGE_TOP_2 - 100),
        PDFItem(logo, 270, PAGE_TOP_2 - 4.5 * cm - 100),
        PDFItem(student_name, MARGIN_LEFT + 250, PAGE_TOP_2 - (4 * cm + 120)),
        PDFItem(stu_pic, PAGE_WIDTH - 220, PAGE_TOP_2 - 120 - 4 * cm),

        PDFItem(table, MARGIN_LEFT + 110, PAGE_TOP_2 - 330),

        PDFItem(stf, MARGIN_LEFT + 350, PAGE_TOP_2 - 380),

        PDFItem(desc, MARGIN_LEFT + 230, PAGE_TOP_2 - 400),

        PDFItem(table1, MARGIN_LEFT + 160, 100),

        PDFItem(tscore, PAGE_WIDTH - 327, 80),
    )
    Page1.add(data)
    return Page1


def get_page_2(stu_details: Tuple[str, str], stu_record, median=0, mode=0, percentile=0, mean=0, avg_accuracy=0,
               avg_attempts=0, total=0, rest_accuracy=0, rest_score={}, att_group=0):
    Page = PDFPage()

    attempt_per = stu_record['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect']).reindex(['Correct', 'Incorrect'], fill_value=0).sum() * 100

    accuracy_per = get_percent_of_attempted_questions(stu_record, 'Outcome (Correct/Incorrect/Not Attempted)', 'Correct'
                                                      , 'Outcome (Correct/Incorrect/Not Attempted)',
                                                      ['Incorrect', 'Correct'])

    background = Image(f"{REQUIREDFILES}/back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    desc1 = Paragraph(text=f"Section 2 ", style=HEADING)
    desc = Paragraph(
        text=f"This section describes {stu_details[0]}'s performance v/s the Rest of the World in Grade {stu_details[2]}.")
    report_table_data = [
        ("Question\nNo.", 'Attempt Status', f"{stu_details[0]}'s\nChoice", 'Correct \nAnswer',
         'Outcome', f"{stu_details[0]}'s\nScore", "% of students\nacross the world\nwho attempted\nthis question",
         "% of students (from\nthose who attempted\nthis ) who got it\ncorrect",
         "% of students\n(from those who\nattempted this)\n"
         "who got it\nincorrect",
         f"World Average\nin this question",),
    ]
    for _, row in stu_record.iterrows():
        report_table_data.append(
            (row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
             row['Outcome (Correct/Incorrect/Not Attempted)'], row['Your score'],
             f"{att_group[row['Question No.']].reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct','Incorrect']].sum() * 100:.2f}%",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Correct'] * 100 :.2f}%",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Incorrect'] *100 :.2f}%",
             f"{rest_score[row['Question No.']]:.2f}")
        )
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table = Table(report_table_data, style=style)
    global FINAL_PERCENTILE
    FINAL_PERCENTILE = percentile
    dets = Paragraph(
        text=f"<b>{stu_details[0]}</b>'s overall percentile in the world is <b>{percentile:.2f}%ile</b>. This indicates "
             f"that {stu_details[0]} has scored more than <b>{percentile:.2f}%</b> of students"
             f" in the World and lesser than <b>{100 - percentile:.2f}%</b> of students<br/>in the world.")
    if isinstance(accuracy_per, Series):
        accuracy_per = accuracy_per.reindex(['Correct', 'Incorrect'], fill_value=0)['Correct']
    tbl_data = (
        (Paragraph(text=f"<b>Overview</b>"),),
        ("Average score of  all\nstudents across the World", Paragraph(text=f"<b>{mean:.2f}</b>"), '',
         f"{stu_details[0]}'s attempts\n(Attempts x 100 / Total\nQuestions)", Paragraph(text=f"<b>{attempt_per:.2f}%</b>"), "",
         f"{stu_details[0]}'s Accuracy\n( Corrects x 100 /Attempts )", Paragraph(text=f"<b>{accuracy_per:.2f}%</b>")),

        (f"Median score of all \nstudents  across the World", Paragraph(text=f"<b>{median:.2f}</b>"), "",
         "Average attempts of all\nstudents across the World", Paragraph(text=f"<b>{avg_attempts.get('Correct', 0) * 100:.2f}%</b>"), f"",
         "Average accuracy of  all\nstudents across the World", Paragraph(text=f"<b>{avg_accuracy.get('Correct', 0) * 100:.2f}%</b>")),

        ("Mode score of all students across\nWorld", Paragraph(text=f"<b>{mode[0]:.2f}</b>"), ''),
    )
    det_style = (
        ('ALIGN', (1, 2), (-1, 2), 'LEFT'),
        ('ALIGN', (1, 4), (-1, 4), 'LEFT'),
        ('ALIGN', (1, 6), (-1, 6), 'LEFT'),
        ('INNERGRID', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 1), (-1, -2), 1, (0, 0, 0)),
        ('BOX', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('BOX', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('BOX', (-2, 1), (-1, -2), 1, (0, 0, 0)),
    )
    tbl = Table(tbl_data, colWidths=(170, 60, cm, 170, 60, cm, 170, 60), style=det_style)
    bar1_img = get_hist(data=[attempt_per.sum(), avg_attempts.reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct','Incorrect']].sum() * 100], xlbl=[f'{stu_details[0]}', 'World'],
                        ylbl="Attempts (%)",
                        title='Comparison of Attempts (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar1 = Image(bar1_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar2_img = get_hist(data=[accuracy_per, avg_accuracy.get('Correct', 0) * 100], xlbl=[f'{stu_details[0]}', 'World'], ylbl="Accuracy (%)",
                        title='Comparison of Accuracy (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar2 = Image(bar2_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    bar3_img = get_hist(data=[total, mean, median, mode[0]], xlbl=[f"{stu_details[0]}", 'Average', 'Median', 'Mode'],
                        ylbl="Score", title='Comparision of Scores', width=5, height=4, label=(0, 3, 6, 9),
                        threshold=0.3, left=0.135, right=0.9, fontsize=15)
    bar3 = Image(bar3_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    data = (PDFItem(background, 0, 0),
            PDFItem(desc1, MARGIN_LEFT + 350, PAGE_TOP_2 - 10),
            PDFItem(desc, MARGIN_LEFT + 200, PAGE_TOP_2 - 30),
            PDFItem(table, MARGIN_LEFT, 450),
            PDFItem(dets, MARGIN_LEFT, 420),
            PDFItem(tbl, MARGIN_LEFT + 10, 280),
            PDFItem(bar3, MARGIN_LEFT + 30, 30),
            PDFItem(bar1, MARGIN_LEFT + PAGE_WIDTH // 3, 30),
            PDFItem(bar2, MARGIN_LEFT + (PAGE_WIDTH // 3) * 2, 30),
            )
    Page.add(data)
    return Page


def get_page_3(stu_details: Tuple[str, str], stu_record, median=0, mode=0, percentile=0, mean=0, avg_accuracy=0,
               avg_attempts=0, total=0, rest_accuracy=0, rest_score={}, att_group=0):
    Page = PDFPage()
    attempt_per = stu_record['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect']).reindex(['Correct', 'Incorrect'], fill_value=0).sum() * 100
    accuracy_per = get_percent_of_attempted_questions(stu_record, 'Outcome (Correct/Incorrect/Not Attempted)', 'Correct'
                                                      , 'Outcome (Correct/Incorrect/Not Attempted)',
                                                      ['Incorrect', 'Correct'])
    background = Image(f"{REQUIREDFILES}/back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    desc1 = Paragraph(text=f"Section 3 ", style=HEADING)
    desc = Paragraph(
        text=f"This section describes {stu_details[0]}'s performance v/s the Rest of {stu_details[1]} in Grade {stu_details[2]}.")
    report_table_data = [
        ("Question\nNo.", "Attempt Status", f"{stu_details[0]}'s\nChoice", 'Correct \nAnswer',
         'Outcome', f"{stu_details[0]}'s\nScore", f"% of students\nacross the {stu_details[1]}"
                                                                            "\nwho attempted\nthis question",
         "% of students (from\nthose who attempted\nthis ) who got it\ncorrect",
         "% of students\n(from those who\nattempted this)\n"
         "who got it\nincorrect",
         f"Average of {stu_details[1]}\nin this question",),
    ]
    for _, row in stu_record.iterrows():
        # breakpoint()
        report_table_data.append(
            (row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
             row['Outcome (Correct/Incorrect/Not Attempted)'], row['Your score'],
             f"{att_group[row['Question No.']].reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct', 'Incorrect']].sum():.2f}%",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Correct'] * 100 :.2f}%",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Incorrect'] * 100 : .2f}%",
             f"{rest_score[row['Question No.']]:.2f}")
        )
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table = Table(report_table_data, style=style)
    dets = Paragraph(
        text=f"<b>{stu_details[0]}</b>'s overall percentile in {stu_details[1]} is <b>{percentile:.2f}%ile</b>. This indicates "
             f"that {stu_details[0]} has scored more than <b>{percentile:.2f}%</b> of students"
             f" in {stu_details[1]} and lesser than <b>{100 - percentile:.2f}%</b> of students in {stu_details[1]}.")
    if isinstance(accuracy_per, Series):
        accuracy_per = accuracy_per.reindex(['Correct', 'Incorrect'], fill_value=0)['Correct']

    tbl_data = (
        (Paragraph(text=f"<b>Overview</b>"),),
        (f"Average score of  all\nstudents in {stu_details[1]}", Paragraph(text=f"<b>{mean:.2f}</b>"), '',
         f"{stu_details[0]}'s attempts\n(Attempts x 100 / Total\nQuestions)", Paragraph(text=f"<b>{attempt_per:.2f}%</b>"), "",
         f"{stu_details[0]}'s Accuracy\n( Corrects x 100 /Attempts )", Paragraph(text=f"<b>{accuracy_per:.2f}%</b>")),

        (f"Median score of all\nstudents  in {stu_details[1]}", Paragraph(text=f"<b>{median:.2f}</b>"), "",
         f"Average attempts of all\nstudents in {stu_details[1]}", Paragraph(text=f"<b>{avg_attempts.get('Correct', 0) * 100:.2f}%</b>"),
         '',
         f"Average accuracy of all\nstudents in {stu_details[1]}", Paragraph(text=f"<b>{avg_accuracy.get('Correct', 0) * 100:.2f}%</b>")),

        (f"Mode score of all students in\n{stu_details[1]}", Paragraph(text=f"<b>{mode:.2f}</b>"), ''),
    )
    det_style = (
        ('ALIGN', (1, 2), (-1, 2), 'LEFT'),
        ('ALIGN', (1, 4), (-1, 4), 'LEFT'),
        ('ALIGN', (1, 6), (-1, 6), 'LEFT'),
        ('INNERGRID', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 1), (-1, -2), 1, (0, 0, 0)),
        ('BOX', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('BOX', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('BOX', (-2, 1), (-1, -2), 1, (0, 0, 0)),
    )
    tbl = Table(tbl_data, colWidths=(170, 60, cm, 170, 60, cm, 170, 60), style=det_style)
    bar1_img = get_hist(data=[attempt_per, avg_attempts.reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct','Incorrect']].sum() * 100], xlbl=[f'{stu_details[0]}', f"{stu_details[1]}"],
                        ylbl="Attempts (%)",
                        title='Comparison of Attempts (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar1 = Image(bar1_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar2_img = get_hist(data=[accuracy_per, avg_accuracy.get('Correct', 0) * 100], xlbl=[f'{stu_details[0]}', f"{stu_details[1]}"],
                        ylbl="Accuracy (%)",
                        title='Comparison of Accuracy (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar2 = Image(bar2_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar3_img = get_hist(data=[total, mean, median, mode], xlbl=[f"{stu_details[0]}", 'Average', 'Median', 'Mode'],
                        ylbl="Score", title='Comparision of Scores', width=5, height=4, label=(0, 3, 6, 9),
                        threshold=0.3, left=0.135, right=0.9, fontsize=15)
    bar3 = Image(bar3_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    data = (PDFItem(background, 0, 0),
            PDFItem(desc1, MARGIN_LEFT + 350, PAGE_TOP_2 - 10),
            PDFItem(desc, MARGIN_LEFT + 195, PAGE_TOP_2 - 30),
            PDFItem(table, MARGIN_LEFT, 450),
            PDFItem(dets, MARGIN_LEFT, 420),
            PDFItem(tbl, MARGIN_LEFT + 10, 280),
            PDFItem(bar3, MARGIN_LEFT + 30, 30),
            PDFItem(bar1, MARGIN_LEFT + PAGE_WIDTH // 3, 30),
            PDFItem(bar2, MARGIN_LEFT + (PAGE_WIDTH // 3) * 2, 30),
            )
    Page.add(data)
    return Page


def get_page_4(stu_details: Tuple[str, str], stu_record, median=0, mode=0, percentile=0, mean=0, avg_accuracy=0,
               avg_attempts=0, total=0, rest_accuracy=0, rest_score={}, att_group=0):
    Page = PDFPage()
    attempt_per = stu_record['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect']).reindex(['Correct', 'Incorrect'], fill_value=0).sum() * 100
    accuracy_per = get_percent_of_attempted_questions(stu_record, 'Outcome (Correct/Incorrect/Not Attempted)', 'Correct'
                                                      , 'Outcome (Correct/Incorrect/Not Attempted)',
                                                      ['Incorrect', 'Correct'])

    background = Image(f"{REQUIREDFILES}/back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    desc1 = Paragraph(text=f"Section 4 ", style=HEADING)
    desc = Paragraph(
        text=f"This section describes {stu_details[0]}'s performance v/s the Best (Top 10%) of the World in Grade {stu_details[2]}")
    report_table_data = [
        ("Question\nNo.", "Attempt Status", f"{stu_details[0]}'s\nChoice", 'Correct\nAnswer',
         'Outcome', f"{stu_details[0]}'s\nScore", "% of students\nacross the World's"
                                                                            "\nBest who attempted\nthis question",
         "% of students (from\nthose who attempted\nthis ) who got it\ncorrect",
         "% of students\n(from those who\nattempted this)\n"
         "who got it\nincorrect",
         f"Average of\nWorld's Best\nin this\nquestion",),
    ]
    for _, row in stu_record.iterrows():
        report_table_data.append(
            (row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
             row['Outcome (Correct/Incorrect/Not Attempted)'], row['Your score'],
             f"{att_group[row['Question No.']].reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct', 'Incorrect']].sum() * 100:.2f}% ",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Correct'] * 100 :.2f}%",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Incorrect'] * 100 : .2f}%",
             f"{rest_score[row['Question No.']]:.2f}")
        )
        if not rest_accuracy.get((row['Question No.'], 'Incorrect'), 0):
            pass
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table = Table(report_table_data, style=style)
    if isinstance(accuracy_per, Series):
        accuracy_per = accuracy_per.reindex(['Correct', 'Incorrect'], fill_value=0)['Correct']

    tbl_data = (
        (Paragraph(text=f"<b>Overview</b>"),),
        ("Average score of \nthe World's Best", Paragraph(text=f"<b>{mean:.2f}</b>"), '',
         f"{stu_details[0]}'s attempts\n(Attempts x 100 / Total\nQuestions)", Paragraph(text=f"<b>{attempt_per:.2f}%</b>"), "",
         f"{stu_details[0]}'s Accuracy\n( Corrects x 100 /Attempts )", Paragraph(text=f"<b>{accuracy_per:.2f}%</b>")),

        (f"Median score of the\nWorld's Best", Paragraph(text=f"<b>{median:.2f}</b>"), "",
         "Average attempts of the\nWorld's Best", Paragraph(text=f"<b>{avg_attempts.get('Correct', 0) * 100:.2f}%</b>"), '',
         "Average accuracy of  the\nWorld's Best", Paragraph(text=f"<b>{avg_accuracy.get('Correct', 0) * 100:.2f}%</b>")),

        ("Mode score of the\nWorld's Best", Paragraph(text=f"<b>{90.00:.2f}</b>"), ''),
    )
    det_style = (
        ('ALIGN', (1, 2), (-1, 2), 'LEFT'),
        ('ALIGN', (1, 4), (-1, 4), 'LEFT'),
        ('ALIGN', (1, 6), (-1, 6), 'LEFT'),
        ('INNERGRID', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 1), (-1, -2), 1, (0, 0, 0)),
        ('BOX', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('BOX', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('BOX', (-2, 1), (-1, -2), 1, (0, 0, 0)),
    )
    tbl = Table(tbl_data, colWidths=(170, 60, cm, 170, 60, cm, 170, 60), style=det_style)
    bar1_img = get_hist(data=[attempt_per, avg_attempts.reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct','Incorrect']].sum() * 100], xlbl=[f'{stu_details[0]}', "World's Best"], ylbl="Attempts (%)",
                        title='Comparison of Attempts (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar1 = Image(bar1_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar2_img = get_hist(data=[accuracy_per, avg_accuracy.get('Correct', 0) * 100], xlbl=[f'{stu_details[0]}', "World's Best"], ylbl="Accuracy (%)",
                        title='Comparison of Accuracy (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar2 = Image(bar2_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar3_img = get_hist(data=[total, mean, median, mode[0]], xlbl=[f"{stu_details[0]}", 'Average', 'Median', 'Mode'],
                        ylbl="Score", title='Comparision of Scores', width=5, height=4, label=(0, 3, 6, 9),
                        threshold=0.3, left=0.135, right=0.9, fontsize=15)
    bar3 = Image(bar3_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    data = (PDFItem(background, 0, 0),
            PDFItem(desc1, MARGIN_LEFT + 350, PAGE_TOP_2 - 10),
            PDFItem(desc, MARGIN_LEFT + 185, PAGE_TOP_2 - 30),
            PDFItem(table, MARGIN_LEFT, 450),
            PDFItem(tbl, MARGIN_LEFT + 10, 280),
            PDFItem(bar3, MARGIN_LEFT + 30, 30),
            PDFItem(bar1, MARGIN_LEFT + PAGE_WIDTH // 3, 30),
            PDFItem(bar2, MARGIN_LEFT + (PAGE_WIDTH // 3) * 2, 30),
            )
    Page.add(data)
    return Page


def get_page_5(stu_details: Tuple[str, str], stu_record, median=0, mode=0, percentile=0, mean=0, avg_accuracy=0,
               avg_attempts=0, total=0, rest_accuracy=0, rest_score={}, att_group=0):
    Page = PDFPage()
    attempt_per = stu_record['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect']).reindex(['Correct', 'Incorrect'], fill_value=0).sum() * 100
    accuracy_per = get_percent_of_attempted_questions(stu_record, 'Outcome (Correct/Incorrect/Not Attempted)', 'Correct'
                                                      , 'Outcome (Correct/Incorrect/Not Attempted)',
                                                      ['Incorrect', 'Correct'])

    background = Image(f"{REQUIREDFILES}/back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    desc1 = Paragraph(text=f"Section 5", style=HEADING)
    desc = Paragraph(
        text=f"This section describes {stu_details[0]}'s performance v/s the Best (Top 10%) of {stu_details[1]} in Grade 3  ")
    report_table_data = [
        ("Question\nNo.", "Attempt Status", f"{stu_details[0]}'s\nChoice", 'Correct\nAnswer',
         'Outcome', f"{stu_details[0]}'s\nScore", f"% of students\nin {stu_details[1]}'s Best"
                                                                            "\nwho attempted\nthis question",
         "% of students (from\nthose who attempted\nthis ) who got it\ncorrect",
         "% of students\n(from those who\nattempted this)\n"
         "who got it\nincorrect",
         f"{stu_details[1]}'s Best\nAverage in\nthis question",),
    ]
    for _, row in stu_record.iterrows():
        report_table_data.append(
            (row['Question No.'], row['Attempt Status'], row['What you marked'], row['Correct Answer'],
             row['Outcome (Correct/Incorrect/Not Attempted)'], row['Your score'],
             f"{att_group[row['Question No.']].reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct', 'Incorrect']].sum() * 100:.2f}% ",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Correct'] * 100 :.2f}%",
             f"{rest_accuracy[row['Question No.']].reindex(['Correct', 'Incorrect'], fill_value=0)['Incorrect'] * 100 : .2f}%",
             f"{rest_score[row['Question No.']]:.2f}")
        )
    style = (
        ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
        ('BACKGROUND', (0, 0), (-1, 0), (0, 0, 0)),
        ('FONT', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('TEXTCOLOR', (0, 0), (-1, 0), (1, 1, 1)),
        ('BOX', (0, 0), (-1, -1), 0.25, (0, 0, 0)),
        ('INNERGRID', (0, 0), (-1, -1), 0.25, (0, 0, 0))
    )
    table = Table(report_table_data, style=style)
    if isinstance(accuracy_per, Series):
        accuracy_per = accuracy_per.reindex(['Correct', 'Incorrect'], fill_value=0)['Correct']

    tbl_data = (
        (Paragraph(text=f"<b>Overview</b>"),),
        (f"Average score of\n{stu_details[1]}'s Best", Paragraph(text=f"<b>{mean:.2f}</b>"), '',
         f"{stu_details[0]}'s attempts\n(Attempts x 100 / Total\nQuestions)", Paragraph(text=f"<b>{attempt_per:.2f}%</b>"), "",
         f"{stu_details[0]}'s Accuracy\n( Corrects x 100 /Attempts )", Paragraph(text=f"<b>{accuracy_per:.2f}%</b>")),

        (f"Median score of\n {stu_details[1]}'s Best", Paragraph(text=f"<b>{median:.2f}</b>"), "",
         f"Average attempts of\n {stu_details[1]}'s Best", Paragraph(text=f"<b>{avg_attempts.get('Correct', 0) * 100:.2f}%</b>"), '',
         f"Average accuracy of\n {stu_details[1]}'s Best", Paragraph(text=f"<b>{avg_accuracy.get('Correct', 0) * 100:.2f}%</b>")),

        (f"Mode score of the\n{stu_details[1]}'s Best", Paragraph(text=f"<b>{mode[0]:.2f}</b>"), ''),
    )
    det_style = (
        ('ALIGN', (1, 2), (-1, 2), 'LEFT'),
        ('ALIGN', (1, 4), (-1, 4), 'LEFT'),
        ('ALIGN', (1, 6), (-1, 6), 'LEFT'),
        ('INNERGRID', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('INNERGRID', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('INNERGRID', (-2, 1), (-1, -2), 1, (0, 0, 0)),
        ('BOX', (0, 1), (1, -1), 1, (0, 0, 0)),
        ('BOX', (3, 1), (4, -2), 1, (0, 0, 0)),
        ('BOX', (-2, 1), (-1, -2), 1, (0, 0, 0)),
    )
    tbl = Table(tbl_data, colWidths=(170, 60, cm, 170, 60, cm, 170, 60), style=det_style)
    bar1_img = get_hist(data=[attempt_per, avg_attempts.reindex(['Correct', 'Incorrect', 'Unattempted'], fill_value=0)[['Correct','Incorrect']].sum() * 100], xlbl=[f'{stu_details[0]}', f"{stu_details[1]}'s Best"],
                        ylbl="Attempts (%)",
                        title='Comparison of Attempts (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar1 = Image(bar1_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar2_img = get_hist(data=[accuracy_per, avg_accuracy.get('Correct', 0) * 100], xlbl=[f'{stu_details[0]}', f"{stu_details[1]}'s Best"],
                        ylbl="Accuracy (%)",
                        title='Comparison of Accuracy (%)', width=5, height=4, threshold=0.2, fontsize=15)
    bar2 = Image(bar2_img, width=8 * cm, height=8 * cm, hAlign='LEFT')
    bar3_img = get_hist(data=[total, mean, median, mode[0]], xlbl=[f"{stu_details[0]}", 'Average', 'Median', 'Mode'],
                        ylbl="Score", title='Comparision of Scores', width=5, height=4, label=(0, 3, 6, 9),
                        threshold=0.3, left=0.135, right=0.9, fontsize=15)
    bar3 = Image(bar3_img, width=8 * cm, height=8 * cm, hAlign='LEFT')

    data = (PDFItem(background, 0, 0),
            PDFItem(desc1, MARGIN_LEFT + 350, PAGE_TOP_2 - 10),
            PDFItem(desc, MARGIN_LEFT + 195, PAGE_TOP_2 - 30),
            PDFItem(table, MARGIN_LEFT+10, 450),
            PDFItem(tbl, MARGIN_LEFT + 10, 280),
            PDFItem(bar3, MARGIN_LEFT + 30, 30),
            PDFItem(bar1, MARGIN_LEFT + PAGE_WIDTH // 3, 30),
            PDFItem(bar2, MARGIN_LEFT + (PAGE_WIDTH // 3) * 2, 30),
            )
    Page.add(data)
    return Page


def get_page_6(all_attempts, all_accuracy, all_scores, stu_details, all_modes, all_medians):
    page = PDFPage()
    background = Image(f"{REQUIREDFILES}/back.png", width=PAGE_WIDTH, height=PAGE_HEIGHT + 100, hAlign='center')
    bar1_plot = get_hist(data=all_attempts,
                         xlbl=[f"{stu_details[0]}'s\nAttempts", 'All of\nworld', f'{stu_details[1]}',
                               'Best of\nWorld',
                               f'Best of\n{stu_details[1]}'], ylbl='Attempts (%)',
                         title='Comparison of Attempts as a %',
                         width=5, height=4,
                         label=(-0.2, 2, 4, 6, 8),
                         threshold=0.2, left=0.135, right=0.89, top=0.85,
                         fontsize=10)
    bar1 = Image(bar1_plot, width=8 * cm, height=8 * cm)
    bar2_plot = get_hist(data=all_accuracy,
                         xlbl=[f"{stu_details[0]}'s\nAccuracy", 'All of\nworld', f'{stu_details[1]}', 'Best of\nWorld',
                               f'Best of\n{stu_details[1]}'], ylbl='Accuracy (%)', title='Comparison of Accuracy as a %',
                         width=5, height=4,
                         label=(-0.2, 2, 4, 6, 8),
                         threshold=0.2, left=0.175, right=0.9, top=0.85,
                         fontsize=10)
    bar2 = Image(bar2_plot, width=8 * cm, height=8 * cm)
    bar3_plot = get_hist(data=all_scores,
                         xlbl=[f"{stu_details[0]}'s\nScore", 'All of\nworld', f'{stu_details[1]}', 'Best of\nWorld',
                               f'Best of\n{stu_details[1]}'], ylbl='Score', title='Comparison of Average Scores', width=5,
                         height=4,
                         label=(-0.2, 2, 4, 6, 8),
                         threshold=0.2, left=0.14, right=0.89, top=0.85,
                         fontsize=10)
    bar3 = Image(bar3_plot, width=8 * cm, height=8 * cm)

    bar_4_plot = get_hist(all_modes, xlbl=[f"{stu_details[0]}'s\nScore", "All of\nworld", f"{stu_details[1]}", "Best of\nWorld", f"Best of\n{stu_details[1]}"], ylbl="Mode Score",
                          title="Comparison of Mode Scores", width=5, height=4, label=(-0.2, 2, 4, 6, 8), threshold=0.2, left=0.14,
                          right=0.89, top=0.85, fontsize=10)
    bar4 = Image(bar_4_plot, width=8*cm, height=8*cm)
    bar_5_plot = get_hist(all_medians, xlbl=[f"{stu_details[0]}'s\nScore", "All of\nworld", f"{stu_details[1]}", "Best of\nWorld", f"Best of\n{stu_details[1]}"],
                          ylbl="Median Score",
                          title="Comparison of Median Scores", width=5, height=4, label=(-0.2, 2, 4, 6, 8), threshold=0.2,
                          left=0.14,
                          right=0.89, top=0.85, fontsize=10)
    bar5 = Image(bar_5_plot, width=8 * cm, height=8 * cm)
    passed = False
    if ROUND == 1:
        if all_scores[0] >=50:
            passed = True
            positive_feedback = f"<font size=12><b>Final Result:&nbsp;&nbsp;&nbsp;&nbsp;{stu_details[0]} has cleared Round I and is eligible for Round II.</b></font>"
        else:
            passed = False
    elif ROUND == 2:
        if FINAL_PERCENTILE >= 70:
            positive_feedback = f"<font size=12><b>Final Result:&nbsp;&nbsp;&nbsp;&nbsp;{stu_details[0]} has cleared Round II and is eligible for Round III.</b></font>"
        else:
            passed = False
    else:
        if FINAL_PERCENTILE >= 70:
            positive_feedback = f"<font size=12><b>Final Result:&nbsp;&nbsp;&nbsp;&nbsp;{stu_details[0]} has scored {FINAL_PERCENTILE}%ile worldwide.</b></font>"
        else:
            passed = False
    if passed:
        sal = positive_feedback
    else:
        sal = f"<font size=12><b>Final Result:&nbsp;&nbsp;&nbsp;&nbsp;{stu_details[0]} has not cleared Round {ROUND} .</b></font>"
    result = Paragraph(
        text=sal)
    title = Paragraph(text=f"<font size=14><b>Consolidated Overview</b></font>")
    data = (
        PDFItem(background, 0, 0),
        PDFItem(title, 320, 900),
        PDFItem(bar1, MARGIN_LEFT + 150, 300),
        PDFItem(bar2, MARGIN_LEFT + 400, 300),
        PDFItem(bar3, MARGIN_LEFT+20, 600),
        PDFItem(bar5, MARGIN_LEFT + 270, 600),
        PDFItem(bar4, MARGIN_LEFT + 520, 600),
        PDFItem(result, MARGIN_LEFT+150, 250)
    )
    page.add(data)
    return page




def create_pdf(st_no):
    ALL_ATTEMPTS = []
    ALL_ACCURACY = []
    ALL_SCORE = []
    ALL_MEDIANS = []
    ALL_MODES = []
    stu = static[static['Student No'] == st_no]
    rec = const[const['Student No'] == st_no]
    grpd = const.groupby('Student No')
    TOTAL = rec['Your score'].sum()
    STU_ATT = rec['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Correct', 'Incorrect']).reindex(['Correct', 'Incorrect'], fill_value=0).sum() * 100
    STU_ACC = get_percent_of_attempted_questions(rec, 'Outcome (Correct/Incorrect/Not Attempted)', 'Correct'
                                                      , 'Outcome (Correct/Incorrect/Not Attempted)',
                                                      ['Incorrect', 'Correct'])
    if isinstance(STU_ACC, Series):
        STU_ACC = STU_ACC.reindex(['Incorrect', 'Correct'], fill_value=0)['Correct']
    ALL_MEDIANS.append(TOTAL)
    '''Global calculations without filter'''
    # Attempts grouped by questions
    QUES_ATTEMPTED = const.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(
        ['Incorrect', 'Correct'])
    # Accuracy grouped by questions
    REST_ACCURACY = \
    const[const['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby('Question No.')[
        'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Score grouped by questions
    REST_SCORE = const.groupby('Question No.')['Your score'].sum() / const['Student No'].nunique()

    # Attempts of each student.
    const.groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])
    # Accuracy grouped by student
    const[const['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby('Student No')[
        'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Average attempts
    AVERAGE_ATTEMPTS = const['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])
    # Average accuracy
    AVERAGE_ACCURACY = const[const['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])][
        'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Median
    MEDIAN = const.groupby('Student No')['Your score'].sum().median()
    ALL_MEDIANS.append(MEDIAN)
    # Mode
    MODE = const.groupby('Student No')['Your score'].sum().mode()
    ALL_MODES.append(TOTAL)
    ALL_MODES.append(MODE[0])
    # Average marks for questions
    MEAN = const.groupby('Student No')['Your score'].sum().mean()

    PERCENTILE = percentileofscore(grpd['Your score'].sum(), rec['Your score'].sum(), kind='strict')

    ALL_ATTEMPTS.append(STU_ATT)
    ALL_ACCURACY.append(STU_ACC)
    ALL_SCORE.append(TOTAL)
    pdf = PDF(dest=f"Output/{stu['Full Name of Candidate'].to_string(index=False)} ({stu['Registration'].to_string(index=False)}).pdf", size=(800, PAGE_HEIGHT))
    # print(f"{PICS}/{stu['Student No'].to_string(index=False)}.jpg")

    pdf.add_page(get_page_1(stu, rec, f"{PICS}/{stu['Student No'].to_string(index=False)}.jpg"))
    ALL_ATTEMPTS.append(AVERAGE_ATTEMPTS.reindex(['Incorrect', 'Correct'], fill_value=0).sum() * 100)
    ALL_ACCURACY.append(AVERAGE_ACCURACY['Correct'].sum() * 100)
    ALL_SCORE.append(round(MEAN))

    pdf.add_page(get_page_2(stu_details=(
        stu['First Name of Candidate'].to_string(index=False), 
        stu['Country of Residence'].to_string(index=False),
        stu['Grade'].to_string(index=False)
        ),
                            mean=MEAN,
                            stu_record=rec, median=MEDIAN, mode=MODE,
                            avg_attempts=AVERAGE_ATTEMPTS,
                            percentile=PERCENTILE,
                            avg_accuracy=AVERAGE_ACCURACY, total=TOTAL, rest_accuracy=REST_ACCURACY,
                            rest_score=REST_SCORE, att_group=QUES_ATTEMPTED))

    '''Country calculations without filter'''
    # Attempts grouped by questions
    QUES_ATTEMPTED = {}
    for cont, data in const.groupby('Country of Residence'):
        QUES_ATTEMPTED[cont] = data.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts()
    # Accuracy grouped by questions
    REST_ACCURACY = {}
    for cont, data in const.groupby('Country of Residence'):
        # REST_ACCURACY[cont] = \
        # data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby('Question No.')[
        #     'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')
        # if REST_ACCURACY[cont].count() < Q_NO:
            REST_ACCURACY[cont] = \
            data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct', 'Unattempted'])].groupby('Question No.')[
                'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Score grouped by questions
    REST_SCORE = {}
    for cont, data in const.groupby('Country of Residence'):
        REST_SCORE[cont] = data.groupby('Question No.')['Your score'].sum() / data['Student No'].nunique()

    # # Attempts of each student.
    # k = {}
    # for cont, data in const.groupby('Country of Residence'):
    #     k[cont] = data.groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])

    # # Accuracy grouped by student
    # k = {}
    # for cont, data in const.groupby('Country of Residence'):
    #     k[cont] = data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Average attempts
    AVERAGE_ATTEMPTS = {}
    for cont, data in const.groupby('Country of Residence'):
        AVERAGE_ATTEMPTS[cont] = data['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(
            ['Incorrect', 'Correct'])

    # Average accuracy
    AVERAGE_ACCURACY = {}
    # breakpoint()
    for cont, data in const.groupby('Country of Residence'):
        accuracy_of_attempted = data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])][
            'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')
        # ques_correct_incorrect = data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])][
        #     'Outcome (Correct/Incorrect/Not Attempted)'].count()
        # if ques_correct_incorrect < Q_NO:
        # if True:
        #     # Since there is only single student in a country the value for a question is unknown
        #     accuracy_of_attempted = data[data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct', 'Unattempted'])][
        #     'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')
        AVERAGE_ACCURACY[cont] = accuracy_of_attempted
        

    # Median
    MEDIAN = {}
    for cont, data in const.groupby('Country of Residence'):
        med = data.groupby('Student No')['Your score'].sum().median()
        MEDIAN[cont] = med
    ALL_MEDIANS.append(MEDIAN[stu['Country of Residence'].values[0]])

    # Mode
    MODE = {}
    for cont, data in const.groupby('Country of Residence'):
        MODE[cont] = data.groupby('Student No')['Your score'].sum().mode()
    ALL_MODES.append(MODE[stu['Country of Residence'].values[0]][0])

    # Average marks per question
    MEAN = {}
    for cont, data in const.groupby('Country of Residence'):
        MEAN[cont] = data.groupby('Student No')['Your score'].sum().mean()

    PERCENTILE = {}
    for cont, data in const.groupby('Country of Residence'):
        PERCENTILE[cont] = percentileofscore(data.groupby('Student No')['Your score'].sum(), TOTAL, kind='strict')

    ALL_ATTEMPTS.append(
        AVERAGE_ATTEMPTS[stu['Country of Residence'].to_string(index=False)][['Incorrect', 'Correct']].sum() * 100)
    ALL_ACCURACY.append(AVERAGE_ACCURACY[stu['Country of Residence'].to_string(index=False)]['Correct'].sum() * 100)
    ALL_SCORE.append(round(MEAN[stu['Country of Residence'].to_string(index=False)]))

    pdf.add_page(get_page_3(stu_details=(
        stu['First Name of Candidate'].to_string(index=False), 
        stu['Country of Residence'].to_string(index=False),
        stu['Grade'].to_string(index=False)
        ),
                            mean=MEAN[stu['Country of Residence'].to_string(index=False)],
                            stu_record=rec, median=MEDIAN[stu['Country of Residence'].to_string(index=False)],
                            mode=MODE[stu['Country of Residence'].to_string(index=False)][0],
                            avg_attempts=AVERAGE_ATTEMPTS[stu['Country of Residence'].to_string(index=False)],
                            percentile=PERCENTILE[stu['Country of Residence'].to_string(index=False)],
                            avg_accuracy=AVERAGE_ACCURACY[stu['Country of Residence'].to_string(index=False)],
                            total=TOTAL,
                            rest_accuracy=REST_ACCURACY[stu['Country of Residence'].to_string(index=False)],
                            rest_score=REST_SCORE[stu['Country of Residence'].to_string(index=False)],
                            att_group=QUES_ATTEMPTED[stu['Country of Residence'].to_string(index=False)]))

    k = const.groupby('Student No')

    '''Calculations of top 5 students in world'''
    # List of top 5 students
    world_top_5_students = k.sum().sort_values(by="Your score", ascending=False).head(5).index
    top_5_data = const[const['Student No'].isin(world_top_5_students)]

    # Attempts grouped by questions
    QUES_ATTEMPTED = top_5_data.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(
        'Correct')

    # Accuracy grouped by questions
    REST_ACCURACY = \
    top_5_data[top_5_data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby(
        'Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Rest Score
    REST_SCORE = top_5_data.groupby('Question No.')['Your score'].sum() / top_5_data['Student No'].nunique()
    # # Attempts of each student
    # top_5_data.groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])

    # # Accuracy of each student
    # top_5_data[top_5_data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Average attempts
    AVERAGE_ATTEMPTS = top_5_data['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])

    # Average accuracy
    AVERAGE_ACCURACY = \
    top_5_data[top_5_data['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])][
        'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Median
    MEDIAN = top_5_data.groupby('Student No')['Your score'].sum().median()
    ALL_MEDIANS.append(MEDIAN)
    # Mode
    MODE = top_5_data.groupby('Student No')['Your score'].sum().mode()
    ALL_MODES.append(MODE[0])
    # Average marks for questions
    MEAN = top_5_data.groupby('Student No')['Your score'].sum().mean()

    PERCENTILE = percentileofscore(top_5_data.groupby('Student No')['Your score'].sum(), TOTAL, kind='strict')

    ALL_ATTEMPTS.append(AVERAGE_ATTEMPTS.reindex(['Incorrect', 'Correct'], fill_value=0).sum() * 100)
    ALL_ACCURACY.append(AVERAGE_ACCURACY['Correct'].sum() * 100)
    ALL_SCORE.append(round(MEAN))

    pdf.add_page(get_page_4(stu_details=(
        stu['First Name of Candidate'].to_string(index=False), 
        stu['Country of Residence'].to_string(index=False),
        stu['Grade'].to_string(index=False)
    ),
                            mean=MEAN,
                            stu_record=rec, median=MEDIAN, mode=MODE,
                            avg_attempts=AVERAGE_ATTEMPTS,
                            avg_accuracy=AVERAGE_ACCURACY, total=TOTAL, rest_accuracy=REST_ACCURACY,
                            rest_score=REST_SCORE, att_group=QUES_ATTEMPTED))

    k = const[const['Country of Residence'] == stu['Country of Residence'].to_string(index=False)].groupby('Student No')

    '''Calculations of top 5 students in country'''
    # List of top 5 students
    country_top_5_students = k.sum().sort_values(by="Your score", ascending=False).head(5).index
    top_5_count = const[const['Student No'].isin(country_top_5_students)]

    # Attempts grouped by questions
    QUES_ATTEMPTED = top_5_count.groupby('Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(
        'Correct')

    # Accuracy grouped by questions
    REST_ACCURACY = \
    top_5_count[top_5_count['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct', 'Unattempted'])].groupby(
        'Question No.')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')
    REST_SCORE = top_5_count.groupby('Question No.')['Your score'].sum() / top_5_count['Student No'].nunique()
    # # Attempts of each student
    # top_5_count.groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])

    # # Accuracy of each student
    # top_5_count[top_5_count['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])].groupby('Student No')['Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Average attempts
    AVERAGE_ATTEMPTS = top_5_count['Outcome (Correct/Incorrect/Not Attempted)'].value_counts(['Incorrect', 'Correct'])

    # Average accuracy
    AVERAGE_ACCURACY = \
    top_5_count[top_5_count['Outcome (Correct/Incorrect/Not Attempted)'].isin(['Incorrect', 'Correct'])][
        'Outcome (Correct/Incorrect/Not Attempted)'].value_counts('Correct')

    # Median
    MEDIAN = top_5_count.groupby('Student No')['Your score'].sum().median()
    ALL_MEDIANS.append(MEDIAN)
    # Mode
    MODE = top_5_count.groupby('Student No')['Your score'].sum().mode()
    ALL_MODES.append(MODE[0])
    # Average marks for questions
    MEAN = top_5_count.groupby('Student No')['Your score'].sum().mean()

    ALL_ATTEMPTS.append(AVERAGE_ATTEMPTS.reindex(['Incorrect', 'Correct'], fill_value=0).sum() * 100)
    ALL_ACCURACY.append(AVERAGE_ACCURACY['Correct'].sum() * 100)
    ALL_SCORE.append(round(MEAN))

    pdf.add_page(get_page_5(stu_details=(
    stu['First Name of Candidate'].to_string(index=False), stu['Country of Residence'].to_string(index=False)),
                            mean=MEAN,
                            stu_record=rec, median=MEDIAN, mode=MODE,
                            avg_attempts=AVERAGE_ATTEMPTS,
                            avg_accuracy=AVERAGE_ACCURACY, total=TOTAL, rest_accuracy=REST_ACCURACY,
                            rest_score=REST_SCORE, att_group=QUES_ATTEMPTED))

    pdf.add_page(get_page_6(ALL_ATTEMPTS, ALL_ACCURACY, ALL_SCORE, (
            f"{stu['Full Name of Candidate'].to_string(index=False)}", 
            stu['Country of Residence'].to_string(index=False),
        ), ALL_MODES, ALL_MEDIANS))
    pdf.prepare((800, PAGE_HEIGHT + 100))

if not os.path.exists('./Output'):
    os.makedirs('./Output')

REQUIREDFILES = './RequiredFiles'
if not os.path.exists('./RequiredFiles'):
    handleError("RequiredFiles Directory does not exist!")

for i, student in static.iterrows():
    try:
        create_pdf(student['Student No'])
    except:
        typer.secho(f"Erro while creating {student['Full Name of Candidate']}'s PDF!", fg=typer.colors.BRIGHT_RED)
print(f"Time taken : {perf_counter() - t}")
print("PDF's generated successfully in Output folder")
sleep(5)