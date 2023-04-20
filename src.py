import configparser
import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import time
import requests
import shutil
from twilio.rest import Client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen


from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import openai

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)


#Read config.ini file
config_obj = configparser.ConfigParser()
config_obj.read("configfile.ini")
linkedin_creds = config_obj["linkdin"]
facebook_creds = config_obj["facebook"]
gpt = config_obj["gpt"]



def speak(audio):
	engine.say(audio)
	engine.runAndWait()

def wishMe():
	hour = int(datetime.datetime.now().hour)
	if hour>= 0 and hour<12:
		speak("Good Morning Saurav!")

	elif hour>= 12 and hour<18:
		speak("Good Afternoon Saurav!")

	else:
		speak("Good Evening Saurav")

	assname =("I am your AI Assistant, SpearTon 1 point o, Ready at your service!")
	speak(assname)

def username():
	speak("How can i Help you, Saurav!!")

def takeCommand():
	
	r = sr.Recognizer()
	
	with sr.Microphone() as source:
		
		print("Listening...")
		audio = r.listen(source)

	try:
		print("Recognizing...")
		query = r.recognize_google(audio, language ='en-in')
		print(f"User said: {query}\n")

	except Exception as e:
		print(e)
		print("Unable to Recognize your voice.")
		return "None"
	
	return query

def sendEmail(to, content):
	server = smtplib.SMTP('smtp.gmail.com', 587)
	server.ehlo()
	server.starttls()
	
	# Enable low security in gmail
	server.login('your email id', 'your email password')
	server.sendmail('your email id', to, content)
	server.close()

if __name__ == '__main__':
	clear = lambda: os.system('cls')
	
	# This Function will clean any
	# command before execution of this python file
	clear()
	wishMe()
	username()

	chrome_options = Options()
	chrome_options.add_argument('--no-sandbox')
	chrome_options.add_argument('--disable-dev-shm-usage')

	# Create a web driver instance
	driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options=chrome_options)
	openai.api_key = gpt['api_key']
	
	while True:
		
		query = takeCommand().lower()
		
		# All the commands said by user will be
		# stored here in 'query' and will be
		# converted to lower case for easily
		# recognition of command
		if 'wikipedia' in query:
			speak('Searching Wikipedia...')
			query = query.replace("wikipedia", "")
			results = wikipedia.summary(query, sentences = 3)
			speak("According to Wikipedia")
			print(results)
			speak(results)

		elif 'open youtube' in query:
			speak("Here you go to Youtube\n")
			driver.get("http://www.youtube.com")
			speak('Do you want me to Search Something in YouTube?')
			youtube_query = takeCommand().lower()
			if 'yes' in youtube_query:
				search_bar = driver.find_element(By.XPATH, "//input[@id='search']")
				speak('What do you wanna Search for?')
				query_search = takeCommand().lower()
				search_bar.send_keys(query_search)
				search_button = driver.find_element(By.XPATH, "//button[@id='search-icon-legacy']")
				search_button.click()
				speak(f"Here are the Top Results for {query_search}!!")
				query_video = takeCommand().lower()

				if "first" in query_video or "top result" in query_video:
					# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
					# driver.get(thumbnail)
					thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[1]//a[@id="video-title"]')))
					thumbnail.click()
				elif "second" in query_video:
					# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
					# driver.get(thumbnail)
					thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]')))
					thumbnail.click()
				elif "third" in query_video:
					# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
					# driver.get(thumbnail)
					thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[3]//a[@id="video-title"]')))
					thumbnail.click()
				elif "fourth" in query_video:
					# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
					# driver.get(thumbnail)
					thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[4]//a[@id="video-title"]')))
					thumbnail.click()
				elif "fifth" in query_video:
					# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
					# driver.get(thumbnail)
					thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[5]//a[@id="video-title"]')))
					thumbnail.click()

		elif "search" in query and "youtube" in query:
			search_query = query.replace("search","").replace("youtube","")
			
			speak(f"Searching {search_query} in YouTube!!")
			driver.get(f"https://www.youtube.com/results?search_query={search_query}")

			speak(f"Here are the Top Results for {query}!!")
			query = takeCommand().lower()

			if "first" in query or "top result" in query:
				# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
				# driver.get(thumbnail)
				thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[1]//a[@id="video-title"]')))
				thumbnail.click()
			elif "second" in query:
				# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
				# driver.get(thumbnail)
				thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]')))
				thumbnail.click()
			elif "third" in query:
				# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
				# driver.get(thumbnail)
				thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[3]//a[@id="video-title"]')))
				thumbnail.click()
			elif "fourth" in query:
				# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
				# driver.get(thumbnail)
				thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[4]//a[@id="video-title"]')))
				thumbnail.click()
			elif "fifth" in query:
				# thumbnail = driver.find_element(By.XPATH, '//*[@id="contents"]/ytd-video-renderer[2]//a[@id="video-title"]').click()
				# driver.get(thumbnail)
				thumbnail = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="contents"]/ytd-video-renderer[5]//a[@id="video-title"]')))
				thumbnail.click()
			# query = takeCommand().lower()

		elif 'open google' in query:
			speak("Opening Google\n")

			# Navigate to Google
			driver.get("http://www.google.com")
			speak('Do you want me to Search Something here?')
			decision = takeCommand().lower()
			if "yes" in decision or "ya" in decision:
				speak('What do you wanna Search for?')
				query_search = takeCommand().lower()

				# Find the search input field
				input_tag = driver.find_element(By.NAME,"q")

				# Enter the search query
				input_tag.send_keys(query_search)

				# Submit the search
				input_tag.send_keys(Keys.ENTER)
				speak(f"Here are the Top Results for {query_search}!!")

				# Find the search results
				search_results = driver.find_elements(By.CSS_SELECTOR,'div.g a')  # This selector targets the search result links

				result_open = takeCommand().lower()
				if "first" in result_open or "top result" in result_open:
					# Click on the search result link
					search_results[1].click()
				elif "second" in result_open:
					# Click on the search result link
					search_results[2].click()
				elif "third" in result_open:
					# Click on the search result link
					search_results[3].click()
				elif "last" in result_open:
					# Click on the search result link
					search_results[len(search_results)-1].click()

		elif 'google' in query and 'search' in query:
			speak("Searching Google\n")
			search_string = query.replace('search','').replace('google','')
			# Navigate to Google
			driver.get(f"https://www.google.com/search?q={search_string}")
			decision = takeCommand().lower()

			speak(f"Here are the Top Results for {search_string}!!")

			# Find the search results
			search_results = driver.find_elements(By.CSS_SELECTOR,'div.g a')  # This selector targets the search result links
			result_open = takeCommand().lower()
			if "first" in result_open or "top result" in result_open:
				# Click on the search result link
				search_results[1].click()
			elif "second" in result_open:
				# Click on the search result link
				search_results[2].click()
			elif "third" in result_open:
				# Click on the search result link
				search_results[3].click()
			elif "last" in result_open:
				# Click on the search result link
				search_results[len(search_results)-1].click()

		elif 'open linkedin' in query:
			speak("Opening Linkedin....")
			# Navigate to LinkedIn
			driver.get("https://www.linkedin.com")
			
			# set your parameters for the database connection URI using the keys from the configfile.ini
			email = linkedin_creds["email"]
			password = linkedin_creds["password"]
			
			# Log in to LinkedIn
			email_input = driver.find_element(By.XPATH, '//*[@id="session_key"]')
			email_input.send_keys(email)
			password_input = driver.find_element(By.XPATH, '//*[@id="session_password"]')
			password_input.send_keys(password)
			password_input.send_keys(Keys.ENTER)

			# Wait for the login to complete and page to load
			WebDriverWait(driver, 10).until(EC.url_contains("feed"))
			while True:
				task = takeCommand().lower()
				if "search" in task and "jobs" in task:
					search_bar = driver.find_element(By.XPATH, "//li-icon[@type='job']")
					search_bar.click()

					# Wait for the login to complete and page to load
					WebDriverWait(driver, 10).until(EC.url_contains("job"))

					# Search for jobs using a specific keyword and location
					# Replace 'KEYWORD' and 'LOCATION' with your desired keyword and location
					query_str = task.replace("search","").replace("jobs","").replace(" in "," ").strip().split(" ")

					search_keyword = " ".join(query_str[:len(query_str)-1]).strip()
					search_location = query_str[len(query_str)-1].strip()
					print(search_keyword, search_location)
					speak(f"Searching {search_keyword} jobs in {search_location}....")


					# search_bar = driver.find_element(By.CSS_SELECTOR, "div.jobs-search-box__input--keyword input")
					search_bar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.jobs-search-box__input--keyword input")))
					search_bar.send_keys(search_keyword)

					# location_bar = driver.find_element(By.CSS_SELECTOR, "div.jobs-search-box__input--location input")
					location_bar = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.jobs-search-box__input--location input")))
					location_bar.send_keys(search_location,Keys.ENTER)

					# Targetting the "All Filters" button and Clicking it to get all the Filters
					all_filters = driver.find_element(By.CSS_SELECTOR, "button.search-reusables__all-filters-pill-button")
					all_filters.click()

					# Fetching the Easy Apply Filter and Activating it
					easy_apply = driver.find_elements(By.XPATH, "//li[@class='search-reusables__secondary-filters-filter']")
					easy_apply_checkbox = easy_apply[7].find_element(By.CSS_SELECTOR, "input.artdeco-toggle__button")
					driver.execute_script("arguments[0].click();", easy_apply_checkbox)

					# Getting the New Results
					show_results_btn = driver.find_element(By.CSS_SELECTOR, "button.search-reusables__secondary-filters-show-results-button")
					show_results_btn.click()

					# New Results
					results = driver.find_elements(By.CSS_SELECTOR, "li.jobs-search-results__list-item occludable-update")

				elif "search" in task and "job" not in task:
					# Wait for the login to complete and page to load
					driver.get("https://www.linkedin.com/feed/")
					WebDriverWait(driver, 10).until(EC.url_contains("feed"))
					search_string = task.replace("search","").strip()
					speak(f"Searching for {search_string} in Linkedin....")

					driver.get(f"https://www.linkedin.com/search/results/people/?keywords={search_string}")
					# Search for a person or job

					search_input = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.CSS_SELECTOR, 'input.search-global-typeahead__input' ) )
					)
					search_input.send_keys(search_string, Keys.ENTER)

					time.sleep(3)

					# Search Result - li reusable-search__result-container
					top_result = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.CSS_SELECTOR, 'li.reusable-search__result-container' ) )
					)
					# Open Profile
					top_result.click()

					# Add Connect Button pv-top-card-v2-ctas 
					connect_btn = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.CSS_SELECTOR, 'div.pvs-profile-actions button' ) )
					)
					# Send Connection Request
					connect_btn.click()

					# Add a Note - aria-label="Add a note"
					send_note = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.XPATH, '//button[@aria-label="Add a note"]' ) )
					)
					send_note.click()

					# Custom Connect Message
					message_textbox = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.XPATH, '//textarea[@id="custom-message"]' ) )
					)
					message_textbox

					response = openai.ChatCompletion.create(
					    model="gpt-3.5-turbo",
					    messages=[
					            {"role": "system", "content": "You are a Linkdin professional"},
					            {"role": "user", "content": f"Compose Custom connection invitation message as Saurav to {search_string}"},
					        ]
					)

					result = ''
					for choice in response.choices:
					    result += choice.message.content
					    
					message_textbox.send_keys(result)

					# Send - aria-label="Send now"
					send_button = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.XPATH, '//button[@aria-label="Send now"]' ) )
					)
					send_button.click()

				elif "message" in task:
					person =task.replace("message ", "").strip()
					speak(f"Messaging {person} in Linkedin....")

					# Wait for the login to complete and page to load
					driver.get("https://www.linkedin.com/messaging")
					WebDriverWait(driver, 10).until(EC.url_contains("messaging"))

					header_tab = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( ( By.CSS_SELECTOR, 'ul.global-nav__primary-items' ) )
					)
					message_tab = header_tab.find_element(By.XPATH, '(//li)[4]')
					message_tab.click()		

					search_tab = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( (By.CSS_SELECTOR, "input#search-conversations") )
					)

					search_tab.send_keys(f"{person}", Keys.ENTER)

					# Wait for the search results to load
					time.sleep(2)

					results = driver.find_elements(By.CSS_SELECTOR, "li.msg-conversations-container__pillar")

					for result in results:
					    name = result.find_element(By.CSS_SELECTOR, "h3").text
					    if person.lower() in name.lower():
					        result.find_element(By.CSS_SELECTOR, "a")
					        result.click()
					# Wait for the search results to load
					time.sleep(2)

					# Asking for message Subject
					speak("What would be the subject of this message?")
					subject = takeCommand().lower()


					message_input = WebDriverWait(driver, 10).until(
					    EC.element_to_be_clickable( (By.XPATH, "//div[@role='textbox']") )
					)
					message_input.click()


					response = openai.ChatCompletion.create(
					    model="gpt-3.5-turbo",
					    messages=[
					            {"role": "system", "content": "You are a indian guy"},
					            {"role": "user", "content": f"message as Saurav to {person},{subject} in hinglish"},
					        ]
					)

					result = ''
					for choice in response.choices:
					    result += choice.message.content

					print(result)
					message = result

					message_input.send_keys(message)
					message_input.submit()


				elif "scroll down" in task:
					actions = ActionChains(driver)
					for i in range(20):
					    actions.send_keys(Keys.ARROW_DOWN).perform()

				elif "scroll up" in task:
					actions = ActionChains(driver)
					for i in range(20):
					    actions.send_keys(Keys.ARROW_UP).perform()

				elif 'exit' in task:
					speak("Thanks for giving me your time")
					exit()


		elif 'play music' in query or "play song" in query:
			speak("Here you go with music")
			# music_dir = "G:\\Song"
			music_dir = "C:\\Users\\GAURAV\\Music"
			songs = os.listdir(music_dir)
			print(songs)
			random = os.startfile(os.path.join(music_dir, songs[1]))

		elif 'the time' in query:
			strTime = datetime.datetime.now().strftime("% H:% M:% S")
			speak(f"Sir, the time is {strTime}")

		elif 'send a mail' in query:
			try:
				speak("What should I say?")
				content = takeCommand()
				speak("whome should i send")
				to = input()
				sendEmail(to, content)
				speak("Email has been sent !")
			except Exception as e:
				print(e)
				speak("I am not able to send this email")

		elif 'how are you' in query:
			speak("I am fine, Thank you")
			speak("How are you, Sir")

		elif 'fine' in query or "good" in query:
			speak("It's good to know that your fine")

		elif "change my name to" in query:
			query = query.replace("change my name to", "")
			assname = query

		elif "change name" in query:
			speak("What would you like to call me, Sir ")
			assname = takeCommand()
			speak("Thanks for naming me")

		elif "what's your name" in query or "What is your name" in query:
			speak("My friends call me")
			speak(assname)
			print("My friends call me", assname)

		elif 'exit' in query:
			speak("Thanks for giving me your time")
			exit()

		elif "who made you" in query or "who created you" in query:
			speak("I have been created by Gaurav.")
			
		elif 'joke' in query:
			speak(pyjokes.get_joke())


		elif "who i am" in query:
			speak("If you talk then definitely your human.")

		elif "why you came to world" in query:
			speak("Thanks to Saurav Singh Rautela. further It's a secret")


		elif "who are you" in query:
			speak("I am your virtual assistant created by Gaurav")

		elif 'reason for you' in query:
			speak("I was created as a Minor project by Mister Saurav Singh Rautela ")

		elif 'news' in query:
			
			try:
				jsonObj = urlopen('''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\''')
				data = json.load(jsonObj)
				i = 1
				
				speak('here are some top news from the times of india')
				print('''=============== TIMES OF INDIA ============'''+ '\n')
				
				for item in data['articles']:
					
					print(str(i) + '. ' + item['title'] + '\n')
					print(item['description'] + '\n')
					speak(str(i) + '. ' + item['title'] + '\n')
					i += 1
			except Exception as e:
				
				print(str(e))


		elif "don't listen" in query or "stop listening" in query:
			speak("for how much time you want to stop jarvis from listening commands")
			a = int(takeCommand())
			time.sleep(a)
			print(a)


		elif "camera" in query or "take a photo" in query:
			ec.capture(0, "Spearton Camera ", "img.jpg")

		
		elif "show note" in query:
			speak("Showing Notes")
			file = open("jarvis.txt", "r")
			print(file.read())
			speak(file.read(6))
					
		# NPPR9-FWDCX-D2C8J-H872K-2YT43
		elif "spearton" in query:
			
			wishMe()
			speak("spearton 1 point o in your service Mister")
			speak(assname)

		elif "weather" in query:
			
			# Google Open weather website
			# to get API of Open weather
			api_key = "Api key"
			base_url = "http://api.openweathermap.org / data / 2.5 / weather?"
			speak(" City name ")
			print("City name : ")
			city_name = takeCommand()
			complete_url = base_url + "appid =" + api_key + "&q =" + city_name
			response = requests.get(complete_url)
			x = response.json()
			
			if x["code"] != "404":
				y = x["main"]
				current_temperature = y["temp"]
				current_pressure = y["pressure"]
				current_humidiy = y["humidity"]
				z = x["weather"]
				weather_description = z[0]["description"]
				print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) +"\n description = " +str(weather_description))
			
			else:
				speak(" City Not Found ")
			
		elif "wikipedia" in query:
			webbrowser.open("wikipedia.com")

		elif "Good Morning" in query:
			speak("A warm" +query)
			speak("How are you Mister")
			speak(assname)



		elif "how are you" in query:
			speak("I'm fine, glad you me that")

		elif "i love you" in query:
			speak("It's hard to understand")

		elif "question" in query or "questions" in query or "what" in query or "where" in query or "when" in query:
			speak('What do you wanna Search for?')
			question = takeCommand().lower()

			response = openai.ChatCompletion.create(
			    model="gpt-3.5-turbo",
			    messages=[
			            {"role": "system", "content": "You are a expert AI assistant"},
			            {"role": "user", "content": f"{question}"},
			        ]
			)

			result = ''
			for choice in response.choices:
			    result += choice.message.content
			print(result)
			speak(result)

