import selenium.common.exceptions
from selenium.webdriver.common.by import By
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from enum import Enum

import re
import pickle
import time
import tkinter as tk
import threading
import keyboard


def parse_excel_file(path):
    book = load_workbook(filename=path)
    sheet = book['Философия']
    res_dict = {}
    i = 1
    current_theme = ''
    while i < 450:
        line = str(sheet[f'A{i}'].value)
        if line.startswith('Тема'):
            current_theme = re.search(r"(?<=:).+", line).group().strip()
            res_dict[current_theme] = {}
            i += 2
            continue
        else:
            if re.fullmatch(r"\d\.\d:", line) is None:
                res_dict[current_theme][line] = sheet[f'B{i}'].value
            i += 1
    with open(r'dictionary_with_answers.txt', 'wb') as f:
        pickle.dump(res_dict, f)

    return res_dict


class BehaviourEndFillingAnswers(Enum):
    wait_until_pressed_key = 2
    send = 1
    do_nothing = 0


class AvailableBrowsers(Enum):
    firefox = webdriver.Firefox
    chrome = webdriver.Chrome
    edge = webdriver.Edge


class EventDelegate:
    def __init__(self):
        self.__event_followers = []

    def __iadd__(self, ehandler):
        self.__event_followers.append(ehandler)
        return self

    def __isub__(self, ehandler):
        self.__event_followers.remove(ehandler)
        return self

    def __call__(self, *args, **kwargs):
        for follower in self.__event_followers:
            follower(*args, **kwargs)


class WebAPI:
    def __init__(self, waiting_time):
        self.driver = None
        self.waiting_time = waiting_time
        self.found_page_with_test_event = EventDelegate()
        self.initialized_event = EventDelegate()
        self.end_filling_test_event = EventDelegate()
        self.can_not_filling_test_event = EventDelegate()
        self.event_stop_thread = threading.Event()

    def initialize(self, password, login, auto_filling_mode, link, browser):
        self.driver = browser()
        self.driver.get(link)
        self.driver.implicitly_wait(self.waiting_time)
        if auto_filling_mode:
            self.driver.find_element(By.XPATH,
                                     value='//*[@id="root"]/div/main/div/div[1]/div[2]/form/label[1]/input').send_keys(
                login)
            self.driver.find_element(By.XPATH,
                                     value='//*[@id="root"]/div/main/div/div[1]/div[2]/form/label[2]/input').send_keys(
                password)
            self.driver.find_element(By.XPATH, value='//*[@id="root"]/div/main/div/div[1]/div[2]/form/button').click()
            self.initialized_event()
        else:
            WebDriverWait(self.driver, 60).until(
                ec.presence_of_element_located((By.XPATH, '//div[@class="position-relative"]')))
            self.initialized_event()

    def start_answering(self, answers, duration_sleep):
        self.driver.switch_to.default_content()
        title = self.driver.find_element(By.XPATH,
                                         value="//a[@class='text-primary-500'][contains(text(), 'Тема')]").text
        theme = ' '.join(title.split()[2:])
        theme_answers = answers[theme]
        frame = self.driver.find_element(by=By.XPATH, value="//iframe[@id='unit-iframe']")
        self.driver.switch_to.frame(frame)
        self.driver.implicitly_wait(5)

        question_blocks = self.driver.find_elements(By.XPATH,
                                                    value='//legend[@class="response-fieldset-legend field-group-hd"]')
        if len(question_blocks) == 0:
            self.driver.implicitly_wait(self.waiting_time)
            count_last_questions = len(self.driver.find_elements(By.XPATH, value='//div[@class="problem"]'))
            i = 0
            return_value = []
            answers = theme_answers.items()
            for answer in answers:
                if i >= len(answers) - count_last_questions - 1:
                    return_value.append(answer)
                i += 1
            self.can_not_filling_test_event(return_value)
            return None

        self.driver.implicitly_wait(self.waiting_time)
        for question_block in question_blocks:
            if self.event_stop_thread.is_set():
                break
            question = question_block.text
            answer = theme_answers[question]
            if '___' in question:
                input_field = self.driver.driver.find_element(By.XPATH,
                                                              value=self.__generate_xpath_for_input_field_question(
                                                                  question))
                input_field.sendkeys(answer)
            else:
                button = self.driver.find_element(By.XPATH,
                                                  value=self.__generate_xpath_for_point_question(question, answer))
                button.click()

            time.sleep(duration_sleep)

        self.end_filling_test_event()

    def __generate_xpath_for_point_question(self, question, answer):
        res = "//legend[@class='response-fieldset-legend field-group-hd']"
        key_word = question.split()
        count_iteration = 5 if len(key_word) > 5 else len(key_word)
        for word in key_word[:count_iteration]:
            res += f"[contains(text(), '{word}')]"
        res += "/..//label[contains(@id, 'choice')]"
        key_word = answer.split()
        count_iteration = 5 if len(key_word) > 5 else len(key_word)
        for word in key_word[:count_iteration]:
            res += f"[contains(text(), '{word}')]"
        return res

    def __generate_xpath_for_input_field_question(self, question):
        res = '//p'
        key_word = question.split()
        count_iteration = 5 if len(key_word) > 5 else len(key_word)
        for word in key_word[:count_iteration]:
            res += f"[contains(text(), '{word}')]"
        res += '/..//input'
        return res

    def try_send_answer(self):
        button_next = self.driver.find_element(By.XPATH, '//button[@data-submitting="Отправка"]')
        if button_next.get_attribute('data-should-enable-submit-button') == 'False':
            return False
        else:
            button_next.click()
            return True

    def try_find_block_questions(self):
        t_end = time.time() + 30
        self.driver.implicitly_wait(1)
        self.driver.switch_to.default_content()
        while t_end > time.time():
            try:
                frame = self.driver.find_element(by=By.XPATH, value="//iframe[@id='unit-iframe']")
                self.driver.switch_to.frame(frame)
                self.driver.find_element(By.XPATH, value='//div[@class="problem"]')
            except selenium.common.exceptions.NoSuchElementException:
                self.load_next_page()
            else:
                self.driver.implicitly_wait(self.waiting_time)
                self.found_page_with_test_event()
                return True
        self.driver.switch_to.default_content()
        return False

    def load_next_page(self):
        self.driver.switch_to.default_content()
        self.driver.find_element(By.XPATH, value="//button[@class='next-btn btn btn-link']").click()

    def quit(self):
        if self.driver is not None:
            self.driver.quit()


class App(tk.Tk):
    def __init__(self, web_api, link=''):
        super().__init__()
        self.login_field = None
        self.start_auto_filling_button = None
        self.stop_auto_filling_button = None
        self.password_field = None
        self.duration_between_answering_field = None
        self.link_field = None
        self.browser_drop_down_menu = None
        self.user_browser_StringVar = None
        self.behaviour_drop_down_menu = None
        self.current_behaviour_tk_StringVar = None
        self.console_text_box = None
        self.answers = None
        self.password = ''
        self.login = ''
        self.web_api = web_api
        self.th = None
        self.auto_filling_flag = False
        self.link = link
        self.duration_between_answering = '0'
        self.behaviour_after_end_filling_answers = BehaviourEndFillingAnswers.do_nothing
        self.user_browser = AvailableBrowsers.edge
        self.web_api.initialized_event += self.set_button_start_auto_filling_enable
        self.web_api.end_filling_test_event += self.on_end_filling_answers
        self.web_api.can_not_filling_test_event += self.on_can_not_fill_answers
        self.initialize_gui()

    def load_user_data(self):
        try:
            with open(r'user_data.txt', 'rb') as f:
                data = pickle.load(f)
                self.password = data['password']
                self.login = data['login']
                self.link = data['link']
                self.auto_filling_flag = data['auto_filling_flag']
                self.duration_between_answering = data['duration_between_answering']
                self.behaviour_after_end_filling_answers = data['behaviour_after_end_filling_answers']
                self.user_browser = data['user_browser']
            self.answers = self.try_get_dict_answers()
        except FileNotFoundError:
            self.log("Файл с сохранёнными данными не найден")
        except EOFError:
            self.log("Файл с пользовательскими даннами не в нужном формате")

    def save_user_data(self):
        dict_to_save = {
            'password': self.password_field.get(),
            'login': self.login_field.get(),
            'link': self.link_field.get(),
            'auto_filling_flag': self.auto_filling_flag,
            'duration_between_answering': self.duration_between_answering_field.get(),
            'behaviour_after_end_filling_answers': BehaviourEndFillingAnswers[
                self.current_behaviour_tk_StringVar.get().replace(' ', '_')],
            'user_browser': AvailableBrowsers[self.user_browser_StringVar.get()]
        }
        with open(r'user_data.txt', 'wb') as f:
            pickle.dump(dict_to_save, f)

    def log(self, message):
        self.console_text_box['state'] = tk.NORMAL
        self.console_text_box.insert(tk.END, str(message) + '\n')
        self.console_text_box['state'] = tk.DISABLED

    def initialize_gui(self):
        self.console_text_box = tk.Text(self, height=10, width=60, state=tk.DISABLED)
        self.console_text_box.grid(row=7, column=0, columnspan=3)
        self.log("Логгер")
        self.load_user_data()

        tk.Label(self, text="Ссылка на страничку").grid(row=0)
        tk.Label(self, text="Логин").grid(row=1)
        tk.Label(self, text="Пароль").grid(row=2)
        tk.Label(self, text="Задержка ввода в секундах").grid(row=3)
        tk.Label(self, text="Поведение после заполнения").grid(row=4)
        tk.Label(self, text="Запускаемый браузер").grid(row=6)

        self.current_behaviour_tk_StringVar = tk.StringVar()
        self.current_behaviour_tk_StringVar.set(self.behaviour_after_end_filling_answers.name.replace('_', ' '))
        self.behaviour_drop_down_menu = tk.OptionMenu(self, self.current_behaviour_tk_StringVar,
                                                      *[beh.name.replace('_', ' ') for beh in
                                                        BehaviourEndFillingAnswers])
        self.behaviour_drop_down_menu.grid(row=4, column=1)

        self.user_browser_StringVar = tk.StringVar()
        self.user_browser_StringVar.set(self.user_browser.name)
        self.browser_drop_down_menu = tk.OptionMenu(self, self.user_browser_StringVar,
                                                    *[browser.name for browser in
                                                      AvailableBrowsers])
        self.browser_drop_down_menu.grid(row=6, column=1)

        self.link_field = tk.Entry(self)
        self.link_field.insert(tk.END, self.link)
        self.login_field = tk.Entry(self)
        self.login_field.insert(tk.END, self.login)
        self.password_field = tk.Entry(self, show='*')
        self.password_field.insert(tk.END, self.password)
        self.duration_between_answering_field = tk.Entry(self)
        self.duration_between_answering_field.insert(tk.END, self.duration_between_answering)

        self.link_field.grid(row=0, column=1)
        self.login_field.grid(row=1, column=1)
        self.password_field.grid(row=2, column=1)
        self.duration_between_answering_field.grid(row=3, column=1)

        tk.Button(self,
                  text='Выход',
                  command=self.close_application).grid(row=18,
                                                       column=0,
                                                       sticky=tk.W,
                                                       pady=4)

        auto_filling_var = tk.IntVar()
        auto_filling_var.set(self.auto_filling_flag)
        (tk.Checkbutton(self,
                        text='Авторегистрация',
                        variable=auto_filling_var,
                        onvalue=True,
                        offvalue=False,
                        command=(lambda:
                                 self.checkbutton_auto_filling_password_changed(auto_filling_var))).grid(row=5,
                                                                                                         column=1,
                                                                                                         sticky=tk.W,
                                                                                                         pady=4))
        show_hide_var = tk.IntVar(value=1)
        tk.Checkbutton(self,
                       text='Скрыть/Показать',
                       variable=show_hide_var,
                       onvalue=True,
                       offvalue=False,
                       command=(lambda: self.checkbutton_show_hide_password_changed(show_hide_var))).grid(row=2,
                                                                                                          column=2,
                                                                                                          sticky=tk.W,
                                                                                                          pady=4)

        tk.Button(self, text='Запустить браузер', command=self.run_browser).grid(row=15,
                                                                                 column=1,
                                                                                 pady=4)

        self.start_auto_filling_button = tk.Button(self, text='Начать заполнение', command=self.start_auto_filling,
                                                   state=tk.DISABLED)
        self.start_auto_filling_button.grid(row=16,
                                            column=1,
                                            pady=4)

        self.stop_auto_filling_button = tk.Button(self, text='Прекратить заполнение', command=self.stop_auto_filling,
                                                  state=tk.DISABLED)
        self.stop_auto_filling_button.grid(row=17,
                                           column=1,
                                           pady=4)

        self.mainloop()

    def start_auto_filling(self):
        self.start_auto_filling_button['state'] = tk.DISABLED
        self.stop_auto_filling_button['state'] = tk.NORMAL
        self.web_api.event_stop_thread.clear()
        if self.web_api.try_find_block_questions():
            self.th = threading.Thread(target=self.web_api.start_answering,
                                       args=(self.answers, (int(self.duration_between_answering_field.get())),))
            self.th.start()
        else:
            self.log("Невозможно найти страницу с тестом")

    def stop_auto_filling(self):
        self.stop_auto_filling_button['state'] = tk.DISABLED
        self.start_auto_filling_button['state'] = tk.NORMAL
        self.web_api.event_stop_thread.set()
        self.web_api.driver.switch_to.default_content()

    def checkbutton_auto_filling_password_changed(self, value):
        self.auto_filling_flag = value.get()

    def checkbutton_show_hide_password_changed(self, value):
        if value.get():
            self.password_field.configure(show="*")
        else:
            self.password_field.configure(show="")

    def set_button_start_auto_filling_enable(self):
        self.start_auto_filling_button['state'] = tk.NORMAL

    def on_end_filling_answers(self):
        self.behaviour_after_end_filling_answers = BehaviourEndFillingAnswers[
            self.current_behaviour_tk_StringVar.get().replace(' ', '_')]
        if self.behaviour_after_end_filling_answers is BehaviourEndFillingAnswers.do_nothing:
            self.stop_auto_filling_button['state'] = tk.DISABLED
            self.set_button_start_auto_filling_enable()
        elif self.behaviour_after_end_filling_answers is BehaviourEndFillingAnswers.wait_until_pressed_key:
            threading.Thread(target=self.wait_until_press_enter).start()
        else:
            self.send_answers()

    def on_can_not_fill_answers(self, answers):
        self.log('Невозможно заполнить поля ответов. Ответы для всего модуля:')
        for answer in answers:
            self.log(answer)
        self.stop_auto_filling_button['state'] = tk.DISABLED
        self.start_auto_filling_button['state'] = tk.NORMAL

    def wait_until_press_enter(self):
        while True:  # making a loop
            try:  # used try so that if user pressed other than the given key error will not be shown
                if keyboard.is_pressed('enter'):  # if key 'q' is pressed
                    self.send_answers()
                    break  # finishing the loop
            except:
                continue

    def send_answers(self):
        if self.web_api.try_send_answer():
            self.web_api.load_next_page()
        else:
            self.log("Невозможно отправить ответы")
            self.stop_auto_filling_button['state'] = tk.DISABLED
            self.start_auto_filling_button['state'] = tk.NORMAL

    def run_browser(self):
        if self.th is not None:
            self.web_api.quit()
            self.th.join()
        self.th = threading.Thread(target=self.web_api.initialize, args=(
            (self.password_field.get()), (self.login_field.get()), self.auto_filling_flag,
            (self.link_field.get()), AvailableBrowsers[self.user_browser_StringVar.get()].value,))
        self.th.start()

    def close_application(self):
        self.save_user_data()
        self.web_api.quit()
        self.quit()

    def try_get_dict_answers(self):
        try:
            with open(r'dictionary_with_answers.txt', 'rb') as f:
                return pickle.load(f)
        except FileNotFoundError:
            print('Отсутсвует файл с ответами')
        return None


if __name__ == '__main__':
    web_driver = WebAPI(15)
    app = App(web_driver)
