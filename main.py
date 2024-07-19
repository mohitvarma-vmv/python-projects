from kivy.clock import Clock
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.screenmanager import ScreenManager, NoTransition, Screen
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.core.window import Window
from kivy.uix.image import Image
from kivy.graphics import Color, Rectangle
import pandas as pd
from email.message import EmailMessage
import ssl
import smtplib

class TitleLabel(Label): #TITLE LABEL TO DISPLAY ON THE LOGIN SCREEN
    def __init__(self, **kwargs):
        super(TitleLabel, self).__init__(**kwargs)
        self.text = "Get Started"
        self.font_name = 'C:\Windows\Fonts/bgothl.ttf'
        self.font_size = '100sp'
        self.size_hint_y = None
        self.height =0
        self.bold = True
        self.color = (1, 1, 1, 1)

class InputForm(Screen): #takes student's name, email amd roll.no to log into the exam
    def __init__(self, email_rollno_mapping, screen_manager, exam_button_callback, **kwargs):
        super(InputForm, self).__init__(**kwargs)
        self.orientation = 'vertical'
        self.spacing=100
        self.screen_manager = screen_manager
        self.email_rollno_mapping = email_rollno_mapping
        self.exam_button_callback = exam_button_callback
        self.title_label = TitleLabel(size_hint=(None,0.1),pos_hint={'center_x':0.5, 'y':0.8})
        self.add_widget(self.title_label)
        text_input_background_color = (1,1,1, 1)
        self.name_input = TextInput(hint_text="Enter your name",size_hint=(None,None),size=(900,60),pos_hint={'center_x':0.7, 'center_y':0.58},background_color=text_input_background_color)
        self.email_input = TextInput(hint_text="Enter your email address",size_hint=(None,None),size=(900,60),pos_hint={'center_x': 0.7, 'center_y':0.50},background_color=text_input_background_color)
        self.rollno_input = TextInput(hint_text="Enter your roll number",size_hint=(None,None),size=(900,60),pos_hint={'center_x': 0.7, 'center_y':0.42},background_color=text_input_background_color)
        self.submit_button = Button(text="Submit",size_hint=(None, None), size=(150,80),pos_hint={'center_x': 0.9,'center_y':0.2},background_color=(120,0.25,0.7,1))
        self.submit_button.bind(on_press=self.verify_credentials)
        self.output_label = Label(text="",size_hint=(1,None),height=60,font_size=50,color=(1,0,0,1),pos_hint={'center_y':0.15})
        self.image2 = Image(source="logo.png", size_hint_x=0.1, size_hint_y=0.1,
                            pos_hint={'center_x': 0.30, 'center_y': 0.5})

        self.add_widget(self.image2)
        self.add_widget(self.name_input)
        self.add_widget(self.email_input)
        self.add_widget(self.rollno_input)
        self.add_widget(self.submit_button)
        self.add_widget(self.output_label)

    def verify_credentials(self, instance): #To verify wether entered email and roll.no. are matching as given in the excel sheet
        name = self.name_input.text
        email = self.email_input.text
        rollno = self.rollno_input.text

        print("Entered Email:", email)
        print("Entered Roll No:", rollno)

        if name and email and rollno:
            if email in self.email_rollno_mapping and self.email_rollno_mapping[email] == rollno:
                user_info = f"Name: {name}\nEmail: {email}\nRoll No: {rollno}"
                self.output_label.text = user_info

                # Set the user email for updating marks later
                exam_screen = self.screen_manager.get_screen("exam_screen")
                exam_screen.user_email = email

                self.exam_button_callback()

                self.screen_manager.current = "exam_screen" #calling exam screen after logging in
            else:
                self.output_label.text = "Email and roll number do not match."
        else:
            self.output_label.text = "Please fill in all fields"

class ExamScreen(Screen): #to design exam screen
    def __init__(self, question_file, key_file, excel_file, **kwargs):
        super(ExamScreen, self).__init__(**kwargs)
        with self.canvas.before:
            Color(0.9, 0.9, 1, 1)  # RGBA for light gray background
            self.rect = Rectangle(size=self.size, pos=self.pos)
            self.bind(size=self._update_rect, pos=self._update_rect)
        self.orientation = 'vertical'
        self.question_file = question_file
        self.key_file = key_file
        self.excel_file = excel_file
        self.questions = []
        self.correct_answers = []
        self.selected_options = []
        self.option_buttons = []
        self.user_email = None  # Store the user's email for updating marks

        self.current_question_index = 0
        self.marks = 0
        self.time_left = 1800  # 30 minutes in seconds
        self.timer_event = None

        self.load_questions_and_answers()

    def _update_rect(self, instance, value): #to update the positon and size of canvas of exam screen
        self.rect.pos = self.pos
        self.rect.size = self.size
    def load_questions_and_answers(self): #maps questions and keys in order to present marks
        try:

            with open(self.question_file, 'r',encoding='utf-8') as q_file, open(self.key_file, 'r') as key_file:
                lines = q_file.readlines()
                self.correct_answers = key_file.read().splitlines()

                current_question = None
                current_options = []

                for line in lines:
                    line = line.strip()
                    if line:
                        if not current_question:
                            current_question = line
                        else:
                            current_options.append(line)

                        if len(current_options) == 4:
                            self.questions.append((current_question, current_options))
                            current_question = None
                            current_options = []

            self.correct_answers = [answer if answer else '-1' for answer in self.correct_answers]
            self.selected_options = [-1] * len(self.questions)  # Initialize the selected options list
            self.display_question()
            self.start_timer()

        except FileNotFoundError:
            self.add_widget(Label(text="Error: Files not found"))

    def start_timer(self):
        self.timer_event = Clock.schedule_interval(self.update_timer, 1) #a method to start a timer

    def update_timer(self, dt):#method to edit the timer
        self.time_left -= 1
        if self.time_left <= 0:
            self.time_left = 0
            self.submit_exam(None)
            return False

        minutes, seconds = divmod(self.time_left, 60)
        time_left_str = f"Time Left: {minutes:02}:{seconds:02}"
        self.timer_label.text = time_left_str

    def display_question(self):#designing amd assigning buttons to the screen

        if 0 <= self.current_question_index < len(self.questions):
            question, options = self.questions[self.current_question_index]
            self.clear_widgets()
            self.option_buttons = []  # Clear and reinitialize option buttons list

            # Main layout with horizontal orientation
            main_layout = BoxLayout(orientation='horizontal', padding=10, spacing=10)

            # Left layout for question number buttons
            left_layout = BoxLayout(orientation='vertical',padding=10, size_hint=(0.3, 0.9))

            scroll_view_left = ScrollView(size_hint=(1, 1))
            question_numbers_layout = GridLayout(cols=1, size_hint_y=None, padding=10, spacing=5)
            question_numbers_layout.bind(minimum_height=question_numbers_layout.setter('height'))

            # Add question number buttons
            question_numbers_layout = GridLayout(cols=6,spacing=30, size_hint_y=None)
            question_numbers_layout.bind(minimum_height=question_numbers_layout.setter('height'))
            for i in range(len(self.questions)):
                question_button = Button(
                    text=str(i + 1),
                    size_hint=(None, None),
                    size=(60, 50),
                    background_color=(0, 10, 0, 1) if self.selected_options[i] != -1 else (1, 0, 0.5, 1)
                )
                question_button.bind(on_press=self.goto_question(i))
                question_numbers_layout.add_widget(question_button)

            scroll_view_left.add_widget(question_numbers_layout)
            left_layout.add_widget(scroll_view_left)

            # Right layout for the question, options, and other labels
            right_layout = BoxLayout(orientation='vertical', size_hint=(0.7, 1), padding=10, spacing=10)

            # Add countdown timer label
            self.timer_label = Label(color=(0,0,0,1), font_size='30sp', font_name = 'C:\Windows\Fonts/arlrdbd.ttf', size_hint_y=None, height=50)
            right_layout.add_widget(self.timer_label)

            #add question label to right layout
            question_label = Label(
                text=f"Question {self.current_question_index + 1}: {question}",
                color=(0,0,0,1),
                font_size='25sp',
                halign='left',
                valign='top',
                size_hint_y=0.1
            )
            question_label.bind(size=question_label.setter('text_size'))
            question_label.height = question_label.texture_size[1] + 20
            right_layout.add_widget(question_label)

            options_layout = GridLayout(cols=1, spacing=10, size_hint=(1,0.8))
            options_layout.bind(minimum_height=options_layout.setter('height'))

            for j, option in enumerate(options, start=1):
                option_button = Button(
                    text=option,
                    size_hint_y=None,
                    height=44,
                    background_color=(1, 1, 1, 0.25),
                    color=(0,0,0,1),
                    font_size= '16sp',
                    text_size=(None, None),  # Allow button width to grow with text size
                    width=self.calculate_button_width(option)  # Calculate width based on text size
                )
                option_button.bind(on_press=self.select_option(j, option_button))
                self.option_buttons.append(option_button)
                options_layout.add_widget(option_button)

            scroll_view_right = ScrollView(size_hint=(1, 1))
            scroll_view_right.add_widget(options_layout)
            right_layout.add_widget(scroll_view_right)

            button_layout = BoxLayout(size_hint_y=None, height=50, spacing=10)

            if self.current_question_index > 0:
                prev_button = Button(
                    text="Previous",
                    size_hint=(None, None),
                    size=(150, 50),
                    pos_hint={'center_x': 0.5},
                    background_color=(0.3,0,0.7,0.7)
                )
                prev_button.bind(on_press=self.previous_question)
                button_layout.add_widget(prev_button)

            if self.current_question_index < len(self.questions) - 1:
                next_button = Button(
                    text="Next",
                    size_hint=(None, None),
                    size=(150, 50),
                    pos_hint={'center_x': 0.5},
                    background_color=(0.3,0,0.7,0.7)
                )
                next_button.bind(on_press=self.next_question)
                button_layout.add_widget(next_button)
            clear_button = Button(
                text="Clear",
                size_hint=(None, None),
                color=(0,0,0,1),
                size=(150, 50),
                pos_hint={'center_x': 0.5},
                background_color=(255,255,0,1)
            )
            clear_button.bind(on_press=self.clear_selection)
            button_layout.add_widget(clear_button)
            submit_button = Button(
                text="Submit Exam",
                size_hint=(None, None),
                size=(150, 50),
                pos_hint={'center_x': 0.5},
                background_color=(100,0,0,1)
            )
            submit_button.bind(on_press=self.submit_exam)
            button_layout.add_widget(submit_button)

            right_layout.add_widget(button_layout)

            # Add left and right layouts to the main layout
            main_layout.add_widget(left_layout)
            main_layout.add_widget(right_layout)

            self.add_widget(main_layout)

            # Highlight the selected option if already selected
            if self.current_question_index < len(self.selected_options):
                selected_option = self.selected_options[self.current_question_index]
                if selected_option > 0:
                    self.option_buttons[selected_option - 1].background_color = (0, 1, 0, 1)
                    # Change background color of the selected option

    def calculate_button_width(self, text):
        # This method calculates the required width based on the text
        # Adjust the multiplier if needed to provide enough space
        return max(150, len(text) * 10)
    def goto_question(self, index):
        def on_select(instance):
            self.current_question_index = index
            self.display_question()

        return on_select

    def select_option(self, option_number, option_button):
        def on_select(instance):
            print(f"Selected Option {option_number} for question {self.current_question_index + 1}")
            self.selected_options[self.current_question_index] = option_number

            for button in self.option_buttons:
                button.background_color = (1, 1, 1, 0.25)
                button.color = (0, 0, 0, 1)
            instance.background_color = (0, 1, 0, 1)
            instance.color=(1,1,1,1)

        return on_select

    def clear_selection(self, instance):
        self.selected_options[self.current_question_index] = -1
        self.display_question()

    def next_question(self, instance):
        if 0 <= self.current_question_index < len(self.questions) - 1:
            self.current_question_index += 1
            self.display_question()

    def previous_question(self, instance):
        if self.current_question_index > 0:
            self.current_question_index -= 1
            self.display_question()
    def submit_exam(self, instance):
        if self.timer_event:
            self.timer_event.cancel()
        if not self.questions or not self.correct_answers:
            self.clear_widgets()
            error_label = Label(text="Error: No questions or answers available", font_size='20sp')
            self.add_widget(error_label)
            return

        self.marks = sum(1 for i, selected_option in enumerate(self.selected_options) if 0 <= i < len(self.correct_answers) and selected_option == int(self.correct_answers[i]))

        # Update marks in Excel
        self.update_marks_in_excel()
        # Send the updated Excel sheet via email
        self.send_email_to_teacher()

        self.clear_widgets()
        result_label = Label(text=f"     Exam is Over\nPress 'ESC' to EXIT", color=(0,0,1,1),font_name = 'C:\Windows\Fonts/framd.ttf' ,font_size='100sp')
        self.add_widget(result_label)

        Clock.schedule_once(self.close_app, 10)

    def update_marks_in_excel(self):
        try:
            df = pd.read_excel(self.excel_file)

            # Update 'Marks' column based on user's email
            if self.user_email is not None:
                df.loc[df['Email'] == self.user_email, 'Marks'] = self.marks

                # Save the updated DataFrame to Excel
                df.to_excel(self.excel_file, index=False)

        except FileNotFoundError:
            print("Error storing marks in Excel: Excel file not found.")
        except Exception as e:
            print("Error storing marks in Excel:", e)


    def send_email_to_teacher(self):
        # Teacher's email address
        teacher_email = 'teacher@gmail.com'

        # Subject and body of the email
        subject = 'Updated Exam Marks'
        body = f'The exam marks have been updated for {self.user_email}. Please find the attached Excel sheet.'

        # Attach the updated Excel sheet
        attachment_path = self.excel_file

        # Email configurations
        sender_email = 'get_started@gmail.com'
        sender_password = 'app_password'

        em=EmailMessage()
        em['From']=sender_email
        em['To']=teacher_email
        em['Subject']=subject
        em.set_content(body)

        with open(attachment_path, 'rb') as f:  # attach file
            file_data = f.read()
            file_name = attachment_path.split('/')[-1]
            em.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        context= ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context)as smtp:
            smtp.login(sender_email,sender_password)
            smtp.send_message(em)

    def close_app(self, dt):
        App.get_running_app().stop()

class UserInfoApp(App):
    def __init__(self, excel_file, question_file, key_file, **kwargs):
        super(UserInfoApp, self).__init__(**kwargs)
        self.excel_file = excel_file
        self.question_file = question_file
        self.key_file = key_file
        self.exam_button_enabled = False

    def build(self):
        screen_manager = ScreenManager(transition=NoTransition())

        Window.fullscreen = 'auto'

        login_screen = Screen(name="login_screen")
        email_rollno_mapping = self.load_data_from_excel()

        def enable_exam_button():
            self.exam_button_enabled = True

        input_form = InputForm(email_rollno_mapping, screen_manager, enable_exam_button)
        login_screen.add_widget(input_form)
        screen_manager.add_widget(login_screen)

        exam_screen = ExamScreen(name="exam_screen", question_file=self.question_file, key_file=self.key_file, excel_file=self.excel_file)
        screen_manager.add_widget(exam_screen)

        return screen_manager

    def load_data_from_excel(self):
        email_rollno_mapping = {}

        try:
            df = pd.read_excel(self.excel_file)

            # Print column names for debugging
            print("Column Names in Excel File:", df.columns.tolist())

            # Assuming 'Email' and 'Roll No' are column names in your Excel file
            email_column = 'Email'
            rollno_column = 'Roll No'

            if email_column in df.columns and rollno_column in df.columns:
                email_rollno_mapping = {row[email_column]: row[rollno_column] for _, row in df.iterrows()}
                print("Loaded Email-RollNo Mapping:", email_rollno_mapping)
            else:
                print(f"Error: '{email_column}' or '{rollno_column}' columns not found in Excel file.")

        except FileNotFoundError:
            print("Error: Excel file not found.")
        except Exception as e:
            print("Error loading data from Excel:", e)

        return email_rollno_mapping

if __name__ == '__main__':
    excel_file = "emailandRoll.no.listofStudents.xlsx"
    question_file = "ques_and_options.txt"
    key_file = "key.txt"
    UserInfoApp(excel_file, question_file, key_file).run()

