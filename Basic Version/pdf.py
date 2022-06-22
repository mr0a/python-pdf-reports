# Library Imports
import jinja2, os, time, sys
import pandas as pd
from PIL import Image
import matplotlib.pyplot as plt
from fpdf import FPDF, HTMLMixin
import numpy as np

# User inputs for data
print("Press Enter for default value")
data_file = input("Enter data location (default is ./Required files/Dummy Data.xlsx): ")
if data_file == '':
    data_file = './Required files/Dummy Data.xlsx'
image_loc = input("Enter Image location (default is ./Required files/Pics for assignment): ")
if image_loc == '':
    image_loc = './Required files/Pics for assignment'

# Load excel data and html template
try:
    df = pd.read_excel(data_file, header=[1])
except:
    print("There is some error with reading the data file.")
    time.sleep(5)
    sys.exit()

try:
    templateLoader = jinja2.FileSystemLoader(searchpath='./Required files/')
    templateEnv = jinja2.Environment(loader=templateLoader)
    template = templateEnv.get_template('template.html')
except:
    print("Error in getting the template file from Required files directory")
    time.sleep(5)
    sys.exit()


datas = {}
analysis = {}
try:
    for index in df.index:
        name = df["Full Name "][index]
        if datas.get(name) == None:
            data = [df["Registration Number"][index], df["Grade "][index], df["Name of School "][index], df["Gender"][index], df["Date of Birth "][index].strftime(r'%d/%m/%Y'), df["City of Residence"][index], df["Country of Residence"][index]]
            datas[name] = {
                "info": data,
                "marks": [],
                "total": 0,
                "correct": 0,
                "incorrect": 0,
                "not_attempted": 0,
                "final_result": df["Final result"][index]
            }
        quest_mark = [
            df["Question No."][index], 
            df["What you marked"][index], 
            df["Correct Answer"][index], 
            df["Outcome (Correct/Incorrect/Not Attempted)"][index],
            df["Score if correct"][index],
            df["Your score"][index]
            ]
        analysis[df["Question No."][index]] = analysis.get(df["Question No."][index], {"correct": 0, "incorrect": 0, "not_attempted": 0})

        if df["Score if correct"][index] == df["Your score"][index]:
            datas[name]["total"] = datas[name].get("total", 0) + df["Score if correct"][index]
            datas[name]["correct"] = datas[name].get("correct", 0) + 1
            analysis[df["Question No."][index]]["correct"] += 1

        else:
            datas[name]["incorrect"] = datas[name].get("incorrect", 0) + 1
            analysis[df["Question No."][index]]["incorrect"] += 1

        if df["Outcome (Correct/Incorrect/Not Attempted)"][index] == 'Unattempted':
            quest_mark[1] = ' '
            datas[name]["not_attempted"] = datas[name].get("not_attempted", 0) + 1
            analysis[df["Question No."][index]]["not_attempted"] += 1

        datas[name]["marks"].append(quest_mark)
except:
    print("Some unknown error with the data")
    time.sleep(5)
    sys.exit()

# Graphs 
analysis_array = [list(data.values()) for data in analysis.values()]
np_array = np.array(analysis_array)
correct = np_array[:, 0]
incorrect = np_array[:, 1]
not_attempted = np_array[:, 2]
labels = [f'Q{index+1}' for index in range(len(correct))]

x = np.arange(len(labels))  # the label locations
width = 0.35  # the width of the bars

fig, ax = plt.subplots(figsize=(10,5))
rects1 = ax.bar(x - width/2, correct, width, label='Correct')
rects2 = ax.bar(x + width/2, incorrect, width, label='Inncorrect')
rects3 = ax.bar(x + width*1.5, not_attempted, width, label='Not Attempted')

# Add some text for labels, title and custom x-axis tick labels, etc.
ax.set_ylabel('Scores')
ax.set_xlabel('Questions')
ax.set_title('Questions grouped by Outcome')
ax.set_xticks(x)
ax.set_xticklabels(labels)
ax.legend()

ax.bar_label(rects1, padding=4)
ax.bar_label(rects2, padding=4)
ax.bar_label(rects3, padding=4)

fig.tight_layout()

if not os.path.exists('./Output'):
    os.makedirs('./Output')
if not os.path.exists('./Required files'):
    os.makedirs('./Required files')

plt.savefig('./Required files/comp.jpg')

columns = ["Question No.", "Your Option", "Correct Option", "Outcome", "Score if correct", "Your Score"]


# Pdf of each student
for student in datas:
    ctx = {
        "name": student,
        "info": datas[student]["info"],
        "columns": columns,
        "marks": datas[student]["marks"],
        "total": datas[student]["total"]
        }
    # Pie chart
    labels = 'Correct', 'Incorrect', 'Not Attempted'
    sizes = [datas[student]["correct"], datas[student]["incorrect"], datas[student]["not_attempted"]]

    explode = (0.1, 0, 0)

    fig1, ax1 = plt.subplots()
    ax1.pie(sizes, explode=explode, labels=labels, autopct='%1.1f%%',
            shadow=True, startangle=90)
    ax1.axis('equal')
    plt.savefig('./Required files/pie.png')

    html = template.render(ctx)

    class MyFPDF(FPDF, HTMLMixin):
        pass

    pdf = MyFPDF()
    pdf.add_page()
    try:
        file_name = f'{image_loc}/{student}'
        img_png = Image.open(f'{file_name}.png')
    except:
        print(f"Error in getting {student}'s image file")
        time.sleep(5)
        sys.exit()

    img_png.save(f'{file_name}.jpg')
    pdf.image(f'{file_name}.jpg',120,52, 50, 50)
    try:
        pdf.image(f'./Required files/icon.jpg', 35, 15, 30, 30)
    except:
        print("Error in loading icon file from Required files directory")
    pdf.rect(5, 5, 200, 287, 'D')
    pdf.write_html(html)
    pdf.set_font('Arial', 'I', 16)
    pdf.cell(0, 10, f' Final Result: {datas[student]["final_result"]} ', 1, 1, 'C')
    pdf.image('./Required files/comp.jpg', w=180, h=90)
    pdf.image('./Required files/pie.png', x =40, w = 130, h=100)
    pdf.rect(5, 5, 200, 287, 'D')
    pdf.output(f'./Output/{student}.pdf', 'F')
print("PDF's generated successfully in Output folder")
time.sleep(5)