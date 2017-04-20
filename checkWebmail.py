# -*- coding: utf-8 -*-

"""
Created on Sat Nov 05 13:42:20 2016
@author: Alexander Hamme
"""

from __future__ import print_function
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, WebDriverException, NoSuchElementException
from selenium.webdriver.support import expected_conditions as ec
import operator
import pyautogui
import time


class WebmailChecker:

    DRIVER_WAIT_TIME = 5

    def __init__(self, user, pw):
        self.driver = None
        self.sessionUrl = ''
        self.sessionID = ''
        self.website_url = "https://email.school.edu"   
        self.username = str(user)
        self.password = str(pw)
        self.user_field_selector = 'zLoginField'                    # These elements can be easily found on your email page,
        self.check_page_loaded_element = 'username'                 # if your browser supports Inspect Element
        self.wait_mail_loaded_element = 'zi_search_inputfield'
        self.email_xpath_locator = "//*[contains(@id, 'zli__TV__')]"
        self.email_class_name_locator = 'RowDouble'
        self.print_feedback = True

    def open_new_session(self):
        '''
        Open new chromedriver session and navigate to URL
        :return: None
        '''
        self.driver = webdriver.Chrome(executable_path="C:\Chromedriver\chromedriver.exe")
        self.driver.maximize_window()
        self.sessionUrl = self.driver.command_executor._url
        self.sessionID = self.driver.session_id
        self.driver.get(self.website_url)
        if self.print_feedback:
            print(self.sessionUrl, self.sessionID)

    def login(self):
        '''
        Log in to email using web page element locators
        :return: None
        '''
        if self.username is None or self.password is None:
            raise SystemExit("Username and/or password not specified")
            
        self.driver.implicitly_wait(self.DRIVER_WAIT_TIME)
        
        try:
            WebDriverWait(self.driver, 2 * self.DRIVER_WAIT_TIME).until (
               ec.visibility_of_element_located((By.NAME, self.check_page_loaded_element))
            )
        except TimeoutException:
            raise TimeoutException("Could not find element: {}".format(self.check_page_loaded_element))

        try:
            user_field, pw_field = list(self.driver.find_elements_by_class_name(self.user_field_selector))

        except WebDriverException:

            try:
                user_field = self.driver.find_element_by_name('username')
                pw_field = self.driver.find_element_by_name('password')

            except WebDriverException:
                raise WebDriverException("Username and/or password fields could not be found")

        if self.print_feedback:
            print("Username and password fields found.")

        # Realistic typing effect (for entertainment purposes)
        for i in range(len(self.username)):
            user_field.send_keys(self.username[i])
            time.sleep(.1)

        for i in range(len(self.password)):
            pw_field.send_keys(self.password[i])
            time.sleep(.1)
        
        self.password = None                # delete password
        pw_field.send_keys(Keys.RETURN)     # alternatively, find search box element and submit()
        
        try:
            WebDriverWait(self.driver, self.DRIVER_WAIT_TIME).until (
               ec.presence_of_element_located((By.ID, self.wait_mail_loaded_element))
            )
        except TimeoutException:
            raise TimeoutException("Login attempt failed")
   
    def load_more(self):
        # emails are dynamically loaded, scrolling is an easy way to load more
        time.sleep(0.5)
        self.driver.execute_script("window.scrollBy(0,-1000)")

    def find_emails(self):
        '''
        Collect emails from web page
        :return: list of Selenium web elements
        '''
        
        self.driver.implicitly_wait(self.DRIVER_WAIT_TIME)

        pyautogui.moveTo(350,410)   # position cursor at center of mail messages, so email is scrolled, not whole page

        # E.g. "300 of 1500 messages" vs "1500 messages" when all are loaded
        messages_loaded = self.driver.find_element_by_xpath('//*[@id="TV__right__text"]').text.split(' ')

        if len(messages_loaded) == 4:  # len("X of X messages") == 4
            total_msgs = int(messages_loaded[2])
            msgs = int(messages_loaded[0])

        else:   # len(messages_loaded) == 2
            msgs = total_msgs = int(messages_loaded[0])

        while len(messages_loaded) == 4:   # when all messages loaded, page displays  'XXXX messages' instead of 'XXX of XXXX messages'

            if self.print_feedback:
                print(msgs, 'of', total_msgs, 'loaded')

            if msgs >= total_msgs:
                break
            if msgs + 50 > total_msgs:
                msgs += total_msgs-msgs
            else:
                msgs += 50

            self.load_more()

            messages_loaded = self.driver.find_element_by_xpath('//*[@id="TV__right__text"]').text.split(' ')

            while len(messages_loaded) == 1:  # If messages are still loading, page displays 'Loading'
                time.sleep(0.5)
                messages_loaded = self.driver.find_element_by_xpath('//*[@id="TV__right__text"]').text.split(' ')

        if self.print_feedback:
            print('\nSearching for emails with XPath locator....')

        try:
            emails = self.driver.find_elements_by_xpath(self.email_xpath_locator)  # use Xpath to locate email elements
            '''    Optional visual loading of emails
            for i in range(len(emails)/100):
              sys.stdout.write(("\r %s"%str(messages_loaded)+" of "+str(total_msgs) + 'messages loaded')+(' '*2)+(" \b[%d"%(i*10)+"%] ")+(''*(100-i))+('##'*i+'|'))
              sys.stdout.flush()
              time.sleep(.1)
            '''
        except NoSuchElementException:
            print('XPath locator may no longer be valid', '\nSearching for emails with class_name locator....')

            try:
                emails = list(self.driver.find_elements_by_class_name(self.email_class_name_locator))
            except NoSuchElementException as nse:
                raise nse

        if self.print_feedback:
            print (len(emails), 'emails located.')

        return emails
        
    def find_senders(self, emails_list):
        '''
        Find which people you have the most emails from and which are the most common subjects and email bodies.
        :param emails_list: list of email web elements
        :return: list of senders, list of most common email bodies, list of most common subject lines
        '''

        months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        senders = []
        email_bodies = []
        subject_lines = []

        if self.print_feedback:
            print("Parsing through collected emails...")

        for eml in emails_list:
            subject_and_body = ((eml.text.encode('utf-8').split('\n'))[0])   #  Example: "Re: Reminder about event - Hi ----, ..."

            subject_lines.append(str((subject_and_body.split('-'))[0]))   # Hyphen always separates subject line from preview of email body
            email_bodies.append(str((subject_and_body.split('-'))[1]))

            sender_and_date = ((eml.text.encode('utf-8').split('\n'))[1]) #  "Anonymous Person 8/01/2016"

            if '/' in sender_and_date.split(' ')[-1]:        # separate name from date if date in m/dd/yyyy format
                senders.append(' '.join(sender_and_date.split(' ')[:-1]))

            elif sender_and_date.split(' ')[-2] in months:
                senders.append(' '.join(sender_and_date.split(' ')[:-2]))
            
        return senders, email_bodies, subject_lines

    def count_occurrences(self, senders_list, email_bods_list, subj_lines_list):
        '''
        Count occurrences of each sender / email body/ subject line
        :param senders_list: list of senders
        :param email_bods_list: list of email bodies
        :param subj_lines_list: list of subject lines
        :return: three dictionaries with counts for each sender / email body / subject line
        '''
        
        names = {name: senders_list.count(name) for name in senders_list}

        em_bods = {bod: email_bods_list.count(bod) for bod in email_bods_list}

        subj_lines = {subj: subj_lines_list.count(subj) for subj in subj_lines_list}

        name_occurrences = sorted(names.items(), key=operator.itemgetter(1), reverse=True)
        email_bod_occurrences = sorted(em_bods.items(), key=operator.itemgetter(1), reverse=True)
        subj_line_occurrences = sorted(subj_lines.items(), key=operator.itemgetter(1), reverse=True)

        return name_occurrences, email_bod_occurrences, subj_line_occurrences
    
    def main(self):

        self.open_new_session()

        self.login()

        email_results = list(self.find_emails())
        
        sender_results, email_bod_results, subj_line_results = self.find_senders(email_results)
        
        sender_occurrences, email_bod_occurrences, subj_line_occurrences = self.count_occurrences(
            sender_results, email_bod_results, subj_line_results
        )

        self.close_session()
        
        return sender_occurrences, email_bod_occurrences, subj_line_occurrences

    def close_session(self):
        try:
            self.driver.close()
            self.driver.quit()
        except WebDriverException:
            self.driver.quit()

        if self.print_feedback:
            print("Chromedriver closed successfully.")

