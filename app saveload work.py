import shinyswatch
from shiny import App, Inputs, Outputs, Session, render, ui, reactive
import pandas
from docxtpl import DocxTemplate
import numpy as np
from pathlib import Path
import win32com.client as win32 
import os,sys
import openai
import pickle


# we can deploy the folder together with shiny for python website
#saving with name.csv file that keep the content user generate, next time they to pick it up by using their's name.csv to upadte the content to keep working ,only all the information being generated then processing to render them to the docxtemplate
#GPT API Setting
API_KEY = "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"
openai.api_key = API_KEY
model_id = 'gpt-3.5-turbo'

os.chdir(sys.path[0])

# A card component wrapper. (for download)
def ui_card(title, *args):
    return (
        ui.div(
            {"class": "card mb-4"},
            ui.div(title, class_="card-header"),
            ui.div({"class": "card-body"}, *args),
        ),
    )


app_ui = ui.page_navbar(
    shinyswatch.theme.superhero(),
    ui.nav( 'The Syllabus Build Guide',
        ui.panel_main(
            ui.navset_tab(
                ui.nav(
                    "Introduction Information",       
                        ui.input_text_area('programname','Program Name',' [Insert program name]',width='500px',height='100px'),
                        ui.input_text_area('classnameandnumber','Class Name and Number',"[Course Title and Number]",width='500px',height='100px'),
                        ui.input_text_area('classtime','Class Time',"[Scheduled Meeting Times]",width='500px',height='100px'),
                        ui.input_text_area('credits','Credits','[Number of credits]',width='500px',height='100px'),
                        ui.input_text_area('coursetype','Course Type','[Core course or Elective]',width='500px',height='100px'),                         
                        ui.input_text_area('instructor','Instructor','[Name, title, email address and phone number]',width='500px',height='100px'),
                        ui.input_text_area('officehours','Office Hours','[SPS Policy: Must state date, time and location; may also indicate by appointment]',width='500px',height='100px'),
                        ui.input_text_area('responsepolicy','Response Policy:','[Include a brief statement about your preferred means of communication and when students should expect a response from you. Will you be available 24/7 or during the workweek only? Will you generally respond within 12 or 24 hours?]',width='500px',height='100px'),
                        ui.input_text_area('TA','Facilitator/Teaching Assistant,','[Name, title, email address and phone number]',width='500px',height='100px'),
                        ui.input_text_area('TAofficehours','TA Office Hours','[SPS policy: Must state date, time and location; may also indicate by appointment]',width='500px',height='100px'),
                        ui.input_text_area('TAresponsepolicy','Response Policy:','[Include a brief statement about your preferred means of communication and when students should expect a response from you. Will you be available 24/7 or during the workweek only? Will you generally respond within 12 or 24 hours?]',width='500px',height='100px'),
                        ),
                ui.nav(
                    "Course Overview",       
                        ui.input_text_area('courseoverview1','First Paragraph','''(a)	Provide a stimulating and descriptive overview of the course. Be sure to include:     
                                           i.	 the course’s main topics    
                                           ii.	for whom the course is designed (e.g., for everyone in the program or primarily for those pursuing a special track)''',width='500px',height='200px'),
                        ui.input_text_area('courseoverview2','Second Paragraph','''(b)	Identify the larger programmatic goals that the course serves. Include:
                                            i.	how the course relates to the primary concepts and principles of the discipline
                                            ii.	how the course fits in with the program curriculum
                                            ''',width='500px',height='200px'),
                        ui.input_text_area('courseoverview3','Third Paragraph','''(c)	Course logistics.
                                            Indicate:
                                            i.	whether the course is a required core course or an elective
                                            ii.	whether or not it will be open, space permitting, to cross-registrants from other fields and/or Columbia University programs; if so which ones
                                            iii.	whether specific competencies or prerequisite knowledge or course work in the discipline are required
                                            iv.	Course Modality (Describe delivery modality: e.g., online, on-campus, hybrid/Hy-flex
                                            v.	Duration. Describe whether the course is: Full semester  Block Week, Partial semester, Residencies; Other format: ___________________________________ ]
                                            ''',width='500px',height='400px'),
                                     
                ),
                ui.nav(
                    "Learning Objectives ",       
                        ui.input_text_area('l1','L1','''[Graduate-level learning objectives encompass learning outcomes that require higher-level functioning, critical analysis, and application to professional fields. Such learning objectives will include observable and actionable verbs such as analyze, critique, design, apply, evaluate, etc. Most SPS courses define 4-6 objectives. Consult a one-page primer from Columbia’s Mailman School. See an example of an SPS graduate course syllabus here. SPS Instructional Design team members can also help you with writing objectives aligned with program goals. Please contact the Senior Director of Instructional Design and Curriculum Support, Ariel Fleurimond, af2830@columbia.edu.
                                            These course-level learning objectives should align with programmatic objectives and be: 
                                            •	observable and measurable
                                            •	designed for the level and purpose of the course
                                            •	be focused on the what the learner will do (not what the instructor will teach)
                                            •	labeled L1, L2, etc. and linked to assignments and activities in the appropriate section.]
                                            ''',width='500px',height='200px'),
                        ui.input_text_area('l2','L2',width='500px',height='100px'),
                        ui.input_text_area('l3','L3',width='500px',height='100px'), 
                        ui.input_text_area('l4','L4',width='500px',height='100px'), 
                        ui.input_text_area('l5','L5',width='500px',height='100px'), 
                        ui.input_text_area('l6','L6',width='500px',height='100px'),   
                          
                ),
                ui.nav(
                     "Readings", 
                     {'id' :'readings'},
                          
                        ui.input_select("citation", "Select Your citation Style", {"APA": "APA", "MLA": "MLA", "Chicago": "Chicago"}),
                        ui.input_text_area('citationinfo','Citation information(Replace the information inside[] )','This is a [Book/Website/Video], the Author(s) is [ NULL ], Title of the book/webpage/article is [ NULL ], Year of publication/Date accessed is [ NULL ], Publisher/Title of the journal is [ NULL ], Page numbers is [ NULL ], Other information: [ NULL ]',width='500px',height='100px'),
                        ui.input_action_button('action_send','Send'),
                        ui.output_text('citecomplete', 'Cite result'),
                        ui.input_text_area('books','Books (Copy the citation into this box)','''[Identify required and recommended readings for the course. Required readings should include a balance of graduate-level practitioner texts and primary academic sources (scholarly articles from peer-reviewed journals in the discipline). Texts have sufficient breadth, depth, and currency for the student to learn the subject at a Master's level and achieve the stated course learning objectives. 
                                            Provide full citations (author, publisher, publication year, etc.), using a recognized citation format, such as MLA, APA or Chicago Style format, after consultation with your academic director. Include page numbers, page counts, and media listening/viewing times so that students can assess the reading workload. Indicate to students where they may find the materials (e.g., Canvas folders, library, purchase from vendor, etc.). Include web links where relevant.
                                            ''',width='500px',height='200px'),
                        ui.input_text_area('others','Other Required Readings (Copy the citation into this box)', 'Other Required Readings (available through Canvas course site or web link)',width='500px',height='200px'),
                        ui.input_text_area('webandvideo','Websites and Videos (Copy the citation into this box)',width='500px',height='200px'),
                         

                ),
                ui.nav(
                     "Assignments and Assessments",  
                    {'id' :'assignment'},                                                   
                        ui.input_text_area('writeassignment','Written assignments','''[Describe here and enumerate the major graduate-level assignments of the course. These descriptions should be high-level to afford flexibility in an approved syllabus. Detailed descriptions should be contained in the Canvas course site.Assignments include all required work to be produced by students and evaluated by the instructor, including: 
                        ●	Written assignments (e.g., case analyses, research projects, project plans, reaction papers, essays, designs, op-eds, etc.)
                        ●	Presentations and performances (e.g., role-playing, strategic interactions, leading discussions, client meetings, etc.)
                        ●	Exams (e.g., tests, mid-terms, in-class assessments, final exams, etc.)
                        ●	Practice (e.g., drafts of required written, designed, or performed work, practice sets, etc.)
                        ●	Online Interaction (synchronous or asynchronous, e.g., discussions, posts, threads, chats, etc.) 
                        ●	Participation (assign no more than 15% of the final grade to participation. Consult with your Academic Director as to program-specific participation grading cap) 
                        ●	Other
                        Include statements regarding 1) how assignments help students achieve the stated learning objectives, build skills toward culminating project or exam, and develop competencies that align with the field/discipline, 2) pitch and degree of difficulty for the intended audience, 3) how you will measure students’ progress toward the course goals (formative assessment),  4) specific criteria you will use to evaluate students’ work, and 5) how and when you will provide feedback. Each of these assignments should indicate the learning objectives stated above (L1, L2, etc.). Indicate the grade weight for each assignment and whether the grade is assigned to the individual or to the group/team. Where applicable, please refer students to the Canvas course site for further specificity on assignments.]
                        ''',width='500px',height='500px'),
                        ui.input_text_area('present','Presentations and performances',width='500px',height='100px'),
                        ui.input_text_area('exams','Exams',width='500px',height='100px'),
                        ui.input_text_area('practice','Practice',width='500px',height='100px'),
                        ui.input_text_area('onlineinteraction','Online Interaction ',width='500px',height='100px'),
                        ui.input_text_area('participation','Participation',width='500px',height='100px'),
                        ui.input_text_area('otherassignment','Others',width='500px',height='100px'),
                         
                ),
                ui.nav(
                     "Grading", 
                     ui.div({'id' :'main-content'},
                     ui.input_action_button("btn", "Insert assignment"),
                     
                     )
                ),
                ui.nav(
                     "Course Schedule/Course Calendar",
                     ui.div({'id' :'coursecalendar'},
                     ui.input_action_button("btn2", "Insert a Week"),
                     
                     )                                                
                ),
                ui.nav(
                    "Course Policies",       
                        ui.input_select('attendance', 'Participation and Attendance', {'You are expected to complete all assigned readings, attend all class sessions, and engage with others in online discussions. Your participation will require that you answer questions, defend your point of view, and challenge the point of view of others. If you need to miss a class for any reason, please discuss the absence with me in advance.' : 'You are expected to complete all assigned readings, attend all class sessions, and engage with others in online discussions. Your participation will require that you answer questions, defend your point of view, and challenge the point of view of others. If you need to miss a class for any reason, please discuss the absence with me in advance. ','I expect you to come to class on time and thoroughly prepared. I will keep track of attendance and look forward to an interesting, lively and confidential discussion. If you miss an experience in class, you miss an important learning moment and the class misses your contribution. More than one absence will affect your grade' : 'I expect you to come to class on time and thoroughly prepared. I will keep track of attendance and look forward to an interesting, lively and confidential discussion. If you miss an experience in class, you miss an important learning moment and the class misses your contribution. More than one absence will affect your grade'},width='500px'),
                        ui.input_select('latework', 'Late work', {'There will be no credit granted to any written assignment that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor.':'There will be no credit granted to any written assignment that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor.','Work that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor will be graded down 1/3 of a grade for every day it is late (e.g., from a B+ to a B).':'Work that is not submitted on the due date noted in the course syllabus without advance notice and permission from the instructor will be graded down 1/3 of a grade for every day it is late (e.g., from a B+ to a B).'},width='500px'),
                        ui.input_text_area('citationpolicy','Citation & Submission','''[All written assignments must use standard citation format (e.g., MLA, APA, Chicago), cite sources, and be submitted to the course website (not via email).]''',width='500px',height='100px'),     
                ),
                ui.nav(
                    "School and University Policies and Resources",       
                        ui.input_select('onlineclass', 'Does this course will use online platform?',{'yes':'Yes','no':'No'}),
                        ui.panel_conditional(
                        "input.onlineclass === 'yes' ", ui.input_text_area("online", "Online platforms policy(No need change)",''' Netiquette
                        Online sessions in this course will be offered through Zoom, accessible through Canvas.  A reliable Internet connection and functioning webcam and microphone are required. It is your responsibility to resolve any known technical issues prior to class. Your webcam should remain turned on for the duration of each class, and you should expect to be present the entire time. Avoid distractions and maintain professional etiquette. 
                        Please note: Instructors may use Canvas or Zoom analytics in evaluating your online participation.
                        More guidance can be found at: https://jolt.merlot.org/vol6no1/mintu-wimsatt_0310.htm
                        Netiquette is a way of defining professionalism for collaborations and communication that take place in online environments. Here are some Student Guidelines for this class:
                        ●	Avoid using offensive language or language that is not appropriate for a professional setting.
                        ●	Do not criticize or mock someone’s abilities or skills.
                        ●	Communicate in a way that is clear, accurate and easy for others to understand.
                        ●	Balance collegiality with academic honesty.
                        ●	Keep an open-mind and be willing to express your opinion.
                        ●	Reflect on your statements and how they might impact others.
                        ●	Do not hesitate to ask for feedback.
                        ●	When in doubt, always check with your instructor for clarification.
                        ''',width='500px',height='500px')
                        ),
                        

                ),
                ui.nav(
                    'Download',
                    ui.input_file("file1", "Choose a file to upload:", multiple=True),
                    ui.input_action_button("loaddata", "load data"),ui.HTML('<br>'),ui.HTML('<br>'),                
                    ui.input_action_button("savedata", "save data"),ui.HTML('<br>'),ui.HTML('<br>'),
                    ui.download_button("download1", "Download data"),ui.HTML('<br>'),ui.HTML('<br>'),
                    ui.download_button("download0", "Download Word"),ui.HTML('<br>'),ui.HTML('<br>'),
                    ui.output_text_verbatim('down'), 
                )
            ),
            
        ),
),
title="Columbia University School of Professional Study",
)






def server(input: Inputs, output: Outputs, session: Session):
    
    #Create reactive value to keep btn
    contet = reactive.Value()
    contet2 = reactive.Value()
    
    
    # GPT conversation function
    def ChatGPT_conversation(conversation):
        response = openai.ChatCompletion.create(
            model=model_id,
            messages=conversation
        )
        conversation.append({'role': response.choices[0].message.role, 'content': response.choices[0].message.content})
        return conversation

    
    # Read the Syllabus Template
    doc = DocxTemplate("SPS Syllabus Template.docx")

    
    #ChatGPT citation
    @reactive.event(input.action_send)
    def citepush():
        conversation = []
        prompt = f'Cite this book by {input.citation()} style format for me: {input.citationinfo()}'
        conversation.append({'role': 'user', 'content': prompt})
        conversation = ChatGPT_conversation(conversation)
        response = ('{0}: {1}\n'.format(conversation[-1]['role'].strip(), conversation[-1]['content'].strip()))
        return response
    
    #citation textout
    @output
    @render.text
    def citecomplete():
        return citepush()
    

    
    #reactive dynamically add input (Grade part) 
    @reactive.Effect
    @reactive.event(input.btn)
    def Grade2():
        #giving a text input area to let user counter for us will be better to get delete
        ptn = input.btn()  
        if ptn > 0:
            assignment = ui.input_text(f'name{input.btn()}', f"Assignment{input.btn()}name", value=input.btn())
            text = ui.input_text(f'{input.btn()}', f"Assignment{input.btn()}weight", value=input.btn())
            select = ui.input_select(f'type{input.btn()}','type',{"Individual grade": "individual grade", "Group grade": "Group grade"})
            ui.insert_ui(
                ui.div({"id": f"inserted-assignment{input.btn()}"}, assignment),
                selector="#main-content",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text{input.btn()}"}, text),
                selector="#main-content",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-select{input.btn()}"}, select),
                selector="#main-content",
                where="beforeEnd",
            )
        qqq = str(ptn)
        contet.set(qqq)

    
    #(Grade part)
    def Grade():
        ptn = contet()  
        weight = []
        content = []
        for i in range(int(ptn)): 
            input_id = f'{i+1}'
            assignment = input[f'name{input_id}']() 
            weight = input[input_id]()
            type = input[f'type{input_id}']()
            entry = {'assignment': f'{assignment}', 'weight': f'{weight}%', 'type': f'{type}'}
            content.append(entry)
        return content

    
    #reactive dynamically add input (course calender part)
    @reactive.Effect
    @reactive.event(input.btn2)
    def coursecalender():
        ptn2 = input.btn2()               
        if ptn2 > 0:
            text = ui.input_text(f'date{input.btn2()}', f"Week{input.btn2()}Date", value=input.btn2())
            text1 = ui.input_text(f'topic{input.btn2()}', f"Week{input.btn2()}topic", value='''Course introductions Foundations of … ''',width='500px')
            text2 = ui.input_text(f'reading{input.btn2()}', f"Week{input.btn2()}reading", value='''Title/author Chapters 1–2, pp 105-135(30 pages) Articles x,y,z, pp 24-44 (20 pages) ''',width='500px')
            text3 = ui.input_text(f'assignments{input.btn2()}', f"Week{input.btn2()}assignments", value='''Statement of purpose due 9/15 ''',width='500px')
            ui.insert_ui(
                ui.div({"id": f"inserted-text{input.btn2()}"}, text),
                selector="#coursecalendar",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text1{input.btn2()}"}, text1),
                selector="#coursecalendar",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text2{input.btn2()}"}, text2),
                selector="#coursecalendar",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text3{input.btn2()}"}, text3),
                selector="#coursecalendar",
                where="beforeEnd",
            )
        qqq2 = str(ptn2)
        contet2.set(qqq2)
    
    #(course calender part)
    def Week():
        ptn2 = contet2()
        date = []
        topic = []
        reading = []
        assignments = []
        content2 = []
        for i in range(int(ptn2)): 
            input_id2 = f'{i+1}'
            date = input[f'date{input_id2}']()
            topic = input[f'topic{input_id2}']()
            reading = input[f'reading{input_id2}']()
            assignments = input[f'assignments{input_id2}']()
            entry2 = {'date' : f'Week{i+1}{date}', 'topic': f'{topic}', 'reading': f'{reading}%', 'assignments': f'{assignments}'}
            content2.append(entry2)
        return content2

    
    

    @reactive.Effect
    @reactive.event(input.savedata)
    def savedata():
         # Define a dictionary to store data with names
        Gradedata = str(contet())
        coursecalenderdata = str(contet2())
        
        dynamicinfo = {}

        for i in range(int(contet())):
            
            assignment_name = input[f'name{i+1}']()              
            text_name = input[f'{i+1}']()              
            select_name = input[f'type{i+1}']()  
            
            # Update the dynamicinfo dictionary directly
            dynamicinfo[f'Assignment{i + 1}'] = assignment_name
            dynamicinfo[f'text{i + 1}'] = text_name
            dynamicinfo[f'select{i + 1}'] = select_name
        
        for y in range(int(contet2())):
            
            data_name = input[f'date{y+1}']()  
            topic_name = input[f'topic{y+1}']()  
            reading_name = input[f'reading{y+1}']()  
            assignments_name = input[f'assignments{y+1}']()  
            
            dynamicinfo[f'date{y + 1}'] = data_name
            dynamicinfo[f'topic{y + 1}'] = topic_name
            dynamicinfo[f'reading{y + 1}'] = reading_name
            dynamicinfo[f'assignments{y + 1}'] = assignments_name


        data_dict = {
            'programname' : input.programname(),
            'classnameandnumber' : input.classnameandnumber(),
            'classtime' : input.classtime(),
            'credits' : input.credits(),
            'coursetype' : input.coursetype(),
            'instructor' : input.instructor(),
            'officehours' : input.officehours(),
            'responsepolicy' : input.responsepolicy(),
            'TA' : input.TA(),
            'TAofficehours' : input.TAofficehours(),
            'TAresponsepolicy' : input.TAresponsepolicy(),
            'courseoverview1' : input.courseoverview1(),
            'courseoverview2' : input.courseoverview2(),
            'courseoverview3' : input.courseoverview3(),
            'l1' : input.l1(),
            'l2' : input.l2(),
            'l3' : input.l3(),
            'l4' : input.l4(),
            'l5' : input.l5(),
            'l6' : input.l6(),
            'books' : input.books(),
            'others' : input.others(),
            'webandvideo' : input.webandvideo(),
            'writeassignment' : input.writeassignment(),
            'present' : input.present(),
            'exams' : input.exams(),
            'practice' : input.practice(),
            'onlineinteraction' : input.onlineinteraction(),
            'participation' : input.participation(),
            'otherassignment' : input.otherassignment(),
            'attendance': input.attendance(),
            'latework': input.latework(),
            'citationpolicy': input.citationpolicy(),
            'contet' : Gradedata,
            'contet2' : coursecalenderdata,
               
        }

        data_dict.update(dynamicinfo)
        print(data_dict)

        # File path of the pickle file
        file_path = 'data.pickle'

        # Save the data dictionary to a pickle file
        with open(file_path, 'wb') as file:
            pickle.dump(data_dict, file)

        print(f'Data saved to {file_path}')
    


    @reactive.Effect
    @reactive.event(input.loaddata)
    def loaddata():
        
        if input.file1():
            file_infos = input.file1()
            print(file_infos[0]["datapath"])
            with open(file_infos[0]["datapath"], "rb") as f:
                loaded_data = pickle.load(f)
                print(loaded_data)

        else:
            file_path = 'data.pickle'
            with open(file_path, 'rb') as file:
                loaded_data = pickle.load(file)
        
        ui.update_text(
            "programname",
            value=f"{loaded_data['programname']}",
        ),
        ui.update_text(
            "classnameandnumber",
            value=f"{loaded_data['classnameandnumber']}",
        ),
        ui.update_text(
            "classtime",
            value=f"{loaded_data['classtime']}",
        ),
        ui.update_text(
            "credits",
            value=f"{loaded_data['credits']}",
        ),
        ui.update_text(
            "coursetype",
            value=f"{loaded_data['coursetype']}",
        ),
        ui.update_text(
            "instructor",
            value=f"{loaded_data['instructor']}",
        ),
        ui.update_text(
            "officehours",
            value=f"{loaded_data['officehours']}",
        ),
        ui.update_text(
            "responsepolicy",
            value=f"{loaded_data['responsepolicy']}",
        ),
        ui.update_text(
            "TA",
            value=f"{loaded_data['TA']}",
        ),
        ui.update_text(
            "TAofficehours",
            value=f"{loaded_data['TAofficehours']}",
        ),
        ui.update_text(
            "TAresponsepolicy",
            value=f"{loaded_data['TAresponsepolicy']}",
        ),
        ui.update_text(
            "courseoverview1",
            value=f"{loaded_data['courseoverview1']}",
        ),
        ui.update_text(
            "courseoverview2",
            value=f"{loaded_data['courseoverview2']}",
        ),
        ui.update_text(
            "courseoverview3",
            value=f"{loaded_data['courseoverview3']}",
        ),
        ui.update_text(
            "l1",
            value=f"{loaded_data['l1']}",
        ),
        ui.update_text(
            "l2",
            value=f"{loaded_data['l2']}",
        ),
        ui.update_text(
            "l3",
            value=f"{loaded_data['l3']}",
        ),
        ui.update_text(
            "l4",
            value=f"{loaded_data['l4']}",
        ),
        ui.update_text(
            "l5",
            value=f"{loaded_data['l5']}",
        ),
        ui.update_text(
            "l6",
            value=f"{loaded_data['l6']}",
        ),
        ui.update_text(
            "books",
            value=f"{loaded_data['books']}",
        ),
        ui.update_text(
            "others",
            value=f"{loaded_data['others']}",
        ),
        ui.update_text(
            "webandvideo",
            value=f"{loaded_data['webandvideo']}",
        ),
        ui.update_text(
            "writeassignment",
            value=f"{loaded_data['writeassignment']}",
        ),
        ui.update_text(
            "present",
            value=f"{loaded_data['present']}",
        ),
        ui.update_text(
            "exams",
            value=f"{loaded_data['exams']}",
        ),
        ui.update_text(
            "practice",
            value=f"{loaded_data['practice']}",
        ),
        ui.update_text(
            "onlineinteraction",
            value=f"{loaded_data['onlineinteraction']}",
        ),
        ui.update_text(
            "participation",
            value=f"{loaded_data['participation']}",
        ),
        ui.update_text(
            "otherassignment",
            value=f"{loaded_data['otherassignment']}",
        ),
        ui.update_select(
            "attendance",
            selected=f"{loaded_data['attendance']}",
        ),
        ui.update_select(
            "latework",
            selected=f"{loaded_data['latework']}",
        ),
        ui.update_text(
            "citationpolicy",
            value=f"{loaded_data['citationpolicy']}",
        ),
        
        gradedata = loaded_data['contet']
        coursecalenderdata = loaded_data['contet2']

        contet.set(gradedata)
        contet2.set(coursecalenderdata)
                
        print(f'Loading......')

        for y in range(int(gradedata)):
            assignment = ui.input_text(f'name{y+1}', f"Assignment{y+1}name", value=loaded_data[f'Assignment{y+1}'])
            text = ui.input_text(f'{y+1}', f"Assignment{y+1}weight", value=loaded_data[f'text{y+1}'])
            select = ui.input_select(f'type{y+1}','type',{"Individual grade": "individual grade", "Group grade": "Group grade"},selected = loaded_data[f'select{y+1}'])
            ui.insert_ui(
                ui.div({"id": f"inserted-assignment{y+1}"}, assignment),
                selector="#main-content",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-text{y+1}"}, text),
                selector="#main-content",
                where="beforeEnd",
            )
            ui.insert_ui(
                ui.div({"id": f"inserted-select{y+1}"}, select),
                selector="#main-content",
                where="beforeEnd",
            )

        for i in range(int(coursecalenderdata)):
                text = ui.input_text(f'date{i+1}', f"Week{i+1}Date", value=loaded_data[f'date{i+1}'])
                text1 = ui.input_text(f'topic{i+1}', f"Week{i+1}topic", value=loaded_data[f'topic{i+1}'],width='500px')
                text2 = ui.input_text(f'reading{i+1}', f"Week{i+1}reading", value=loaded_data[f'reading{i+1}'],width='500px')
                text3 = ui.input_text(f'assignments{i+1}', f"Week{i+1}assignments", value=loaded_data[f'assignments{i+1}'],width='500px')
                ui.insert_ui(
                    ui.div({"id": f"inserted-text{i+1}"}, text),
                    selector="#coursecalendar",
                    where="beforeEnd",
                )
                ui.insert_ui(
                    ui.div({"id": f"inserted-text1{i+1}"}, text1),
                    selector="#coursecalendar",
                    where="beforeEnd",
                )
                ui.insert_ui(
                    ui.div({"id": f"inserted-text2{i+1}"}, text2),
                    selector="#coursecalendar",
                    where="beforeEnd",
                )
                ui.insert_ui(
                    ui.div({"id": f"inserted-text3{i+1}"}, text3),
                    selector="#coursecalendar",
                    where="beforeEnd",
                )
        

    
    #render content to the doc
    @output
    @render.text
    def down():
        #Create dynamcic input and bullets for learning objectives
        bullets = [
                        input.l1(),
                        input.l2(),
                        input.l3(),
                        input.l4(),
                    ]
        if input.l5() != '':
                bullets += (input.l5(),)
        if input.l6() != '':
                bullets += (input.l6(),)

        particapation = input.participation()
        if input.otherassignment() != '':
                particapation += f'\n\n7. Others\n{input.otherassignment()}'
            

        if input.onlineclass() == 'no':
            online = ''
        else:
            online = input.online()

        grade = Grade()
        week = Week()
        context = { 
            'programname' : input.programname(),
            'classnameandnumber' : input.classnameandnumber(),
            'classtime' : input.classtime(),
            'credits' : input.credits(),
            'coursetype' : input.coursetype(),
            'instructor' : input.instructor(),
            'officehours' : input.officehours(),
            'responsepolicy' : input.responsepolicy(),
            'TA' : input.TA(),
            'TAofficehours' : input.TAofficehours(),
            'TAresponsepolicy' : input.TAresponsepolicy(),
            'courseoverview1' : input.courseoverview1(),
            'courseoverview2' : input.courseoverview2(),
            'courseoverview3' : input.courseoverview3(),
            'bullets' : bullets,
            'books' : input.books(),
            'others' : input.others(),
            'webandvideo' : input.webandvideo(),
            'writeassignment' : input.writeassignment(),
            'present' : input.present(),
            'exams' : input.exams(),
            'practice' : input.practice(),
            'onlineinteraction' : input.onlineinteraction(),
            'participation' : particapation,
            'tbl_contents': grade,
            'week_table': week,
            'attendance': input.attendance(),
            'latework': input.latework(),
            'citationpolicy': input.citationpolicy(),
            'online': online       
            } 

        doc.render(context)
        doc.save("Complete.docx")

        print('Already send all the information, please wait for processing!')

    
    #doc download
    @session.download()
    def download0():
        # This is the simplest case. The implementation simply returns the path to a
        # file on disk.
        path = Path(__file__).parent / "Complete.docx"
        return str(path)

    @session.download()
    def download1():
        path = Path(__file__).parent / "data.pickle"
        return str(path)





app = App(app_ui, server)