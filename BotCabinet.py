from selenium import webdriver
from collections import OrderedDict
import pandas as pd
import datetime
import math
import time
import sys
import os
import re
from secrets import username, password

import tkinter as tk
import tkinter.font as tkFont
from tkinter import *

"""Updates à faire :
- gérer toutes les écritures des mois proprement (attendre le mois de décembre)
- voir ligne 77 "trouver un moyen de quitter code sans quitter interface graphique"
"""


class BotCabinet:
	def __init__(self, app, remaining_time_label, dossier_list, warning_list, error_list, state_label):
		#correpondance numérique des mois de l'année (plusieurs mois en trop à partir de juillet car je ne connais pas encore le format)
		self.dict_months = OrderedDict({"janv.": 1, "févr.": 2, "mars": 3, "avr.": 4, "mai": 5, "juin": 6, 
			"juil.": 7, "août": 8, "sep.": 9, "sept.": 9, "oct.": 10, "nov.": 11, "déc.": 12, "dec.": 12})

		#TEMPORAIRE : dictionnaire backup ne se basant que sur la 1ère lettre, au cas où le format des mois à partir de juillet ne soit pas prévu
		self.dict_backup = OrderedDict({"j": 7, "a": 8, "s": 9, "o": 10, "n": 11, "d": 12})

		#Date du jour
		self.current_date = datetime.date.today()

		#Open a Chrome page
		self.driver = webdriver.Chrome()

		#Attendre que l'élément de la page web soit trouvé
		self.driver.implicitly_wait(60)

		#On récupère l'interface graphique (Tkinter), ainsi que ses listes et labels, pour pouvoir les mettre à jour
		self.app = app
		self.remaining_time_label = remaining_time_label
		self.dossier_list = dossier_list
		self.warning_list = warning_list
		self.error_list = error_list
		self.state_label = state_label

	def login(self):
		self.driver.get("https://mail.yahoo.com/d/folders/2")

		#Write the login and password in the iframe, then submit
		self.driver.find_element_by_xpath('//*[@id="login-username"]').send_keys(username)
		self.driver.find_element_by_xpath('//*[@id="login-signin"]').click()
		self.driver.find_element_by_xpath('//*[@id="login-passwd"]').send_keys(password)
		self.driver.find_element_by_xpath('//*[@id="login-signin"]').click()

	def get_sent_emails(self, nb_days=15):
		self.remaining_time_label.config(text="Vous avez choisi : {} jour(s)".format(nb_days))
		self.app.update()

		#On commence par ouvrir le tableau des dossiers enregistrés
		df = pd.read_csv("YOUR_PATH\\dossiers.csv", dtype={'dossier': pd.Int64Dtype()}, sep=";").dropna()

		#On remet les dates dans le bon format (datetime.date)
		if not df.empty:
			first_date = df.date.values[0]
			if '-' in first_date:
				dates = [datetime.datetime.strptime(date, '%Y-%m-%d').date() for date in df.date.values]
			elif '/' in first_date:
				dates = [datetime.datetime.strptime(date, '%d/%m/%Y').date() for date in df.date.values]
			else:
				self.state_label.config(text="ERREUR : le format des dates dans le tableau Excel n'est pas reconnu", fg="red")
				self.app.update()
				#Trouver un moyen de quitter le code sans quitter l'interface graphique (pas important car ça ne devrait jamais arriver)
			df.loc[:, 'date'] = dates

		#On fait défiler la page, jusqu'à atteindre tous les mails souhaités
		scrolling = False
		expiration = False
		subject_list_before, date_list_before, receiver_list_before = [], [], []
		while not expiration:
			if scrolling:
				#On scrolle jusqu'au dernier mail actuellement affiché sur la page, pour mettre à jour la liste de mails
				webdriver.ActionChains(self.driver).move_to_element(subject_list[-1]).perform()
			else:
				scrolling = True

			error_len = False
			mail_list_updated = True
			error_loading = False
			while True:
				#Lorsqu'on fait défiler la page, de nouveaux mails s'affichent tandis que d'autres disparaissent
				#On ne traite donc que les nouveaux mails, en soustrayant l'ancienne liste à la nouvelle
				subject_list_after = self.driver.find_elements_by_xpath('//*[@data-test-id="message-subject"]')
				date_list_after = self.driver.find_elements_by_xpath('//time[@role="gridcell"]')
				receiver_list_after = self.driver.find_elements_by_xpath('//*[@data-test-id="senders"]')
				subject_list = [subject for subject in subject_list_after if subject not in subject_list_before]
				date_list = [date for date in date_list_after if date not in date_list_before]
				receiver_list = [receiver for receiver in receiver_list_after if receiver not in receiver_list_before]

				#On vérifie qu'il y a le même nombre de mails, de dates et de destinataires
				len_subject = len(subject_list)
				len_date = len(date_list)
				len_receiver = len(receiver_list)
				if not (len_subject == len_date == len_receiver):
					message = "ERREUR : le nombre d'objets, de dates et de destinataires n'est pas le même : {} / {} / {} -> ATTENDRE".format(len_subject, len_date, len_receiver)
					if self.error_list.get(END) != message:
						self.error_list.insert(END, (message))
						self.state_label.config(text=message, fg="orange")
						self.app.update()
					error_len = True
				#subject_list peut être vide si la liste ne s'est pas mise à jour (cela arrive si l'on réduit la page Yahoo, par exemple)
				elif len_subject == 0:
					message = "ERREUR : les mails ne se mettent pas à jour -> REAFFICHER LA PAGE YAHOO !"
					if self.error_list.get(END) != message:
						self.error_list.insert(END, (message))
						self.state_label.config(text=message, fg="orange")
						self.app.update()
					mail_list_updated = False
				# Le contenu de la page peut "mal" se charger : certains éléments sont bien localisés mais ne contiennent rien
				elif "" in [element.text for element in date_list + receiver_list]:
					message = "ERREUR : le contenu de la page ne s'est pas chargé complètement -> ATTENDRE"
					if self.error_list.get(END) != message:
						self.error_list.insert(END, (message))
						self.state_label.config(text=message, fg="orange")
						self.app.update()
					error_loading = True
				else:
					break
			
			if error_len or (not mail_list_updated) or error_loading:
				self.error_list.insert(END, ("-> RESOLUE(S)"))
				self.app.update()
			subject_list_before = subject_list_after
			date_list_before = date_list_after
			receiver_list_before = receiver_list_after

			for subject, date, receiver in zip(subject_list, date_list, receiver_list):
				subject = subject.text
				date = date.text
				receiver = receiver.text

				#On modifie le format de la date du mail
				if ":" in date:
					date = self.current_date
				elif len(date.split()) == 2:
					day, month = date.split()
					day = int(day)
					#Gérer toutes les écritures des mois
					if month in self.dict_months:
						month = self.dict_months[month]
					elif month[0] in self.dict_backup:
						month = self.dict_backup[month[0]]
					else:
						self.warning_list.insert(END, ("ATTENTION : le format de la date (='{}') du mail '{}' n'est pas détecté".format(subject, date)))
						self.app.update()
						continue
					year = self.current_date.year
					#On gère un problème de date pouvant survenir juste après le passage à une nouvelle année
					date = datetime.date(year, month, day)
					if date > self.current_date:
						date = datetime.date(year-1, month, day)
				elif "/" in date:
					date = datetime.datetime.strptime(date, '%d/%m/%Y').date()
				else:
					self.warning_list.insert(END, ("ATTENTION : le format de la date (='{}') du mail '{}' n'est pas détecté".format(subject, date)))
					self.app.update()
					continue

				#On met à jour la barre de progression du programme
				days_between = (self.current_date - date).days
				progress = round(days_between / nb_days * 100, 1)
				self.state_label.config(text="PROGRESSION : {} %".format(progress), fg="blue")
				self.app.update()

				#Si la date du mail est trop ancienne, on sort de la boucle (on aurait pu se reservir de days_between, mais ca permet de voir une autre façon de faire)
				if date <= self.current_date - datetime.timedelta(nb_days):
					expiration = True
					break

				#On récupère le numéro de dossier à la fin de l'objet du mail
				dossier = subject[-8:]

				if dossier.isdigit() and dossier[:2] in {'19', '20'}:
					dossier = int(dossier)
				else:
					dossier_finded = False
					#Si le numéro de dossier n'est pas à la fin de l'objet, on le cherche dans tout l'objet
					potential_dossiers = re.findall(r"\d+", subject)
					for dossier in potential_dossiers:
						if len(dossier) == 8 and dossier[:2] in {'19', '20'}:
							dossier_finded = True
							dossier = int(dossier)
							break
					if not dossier_finded:
						self.warning_list.insert(END, ("ATTENTION : n° dossier absent du mail '{}' du {}".format(subject, date)))
						self.app.update()
						continue

				#Si le dossier existe, on met à jour sa date de traitement, sinon on l'ajoute au dataframe.
				if dossier in df.dossier.values:
					if date > df.loc[df.dossier == dossier, 'date'].values[0]:
						df.loc[df.dossier == dossier, ['date', 'destinataire', 'objet']] = [date, receiver, subject]
						self.dossier_list.insert(END, ("dossier {} : modifié".format(dossier)))
						self.app.update()
					#Section sous commentaires car on s'en fiche de savoir qu'un dossier est ignoré
					# else:
					# 	print("dossier {} : ignoré".format(dossier))
				else:
					df = df.append({'dossier': dossier, 'date': date, 'destinataire': receiver, 'objet': subject}, ignore_index=True)
					self.dossier_list.insert(END, ("dossier {} : ajouté".format(dossier)))
					self.app.update()

		#On trie le tableau de dossiers par date (des plus anciennes aux plus récentes)
		df.sort_values(by="date", inplace=True)

		#On exporte le tableau (toujours sous le même nom)
		df.to_csv("YOUR_PATH\\dossiers.csv", sep=";", index=False, encoding="utf-8-sig")

		#Pour faciliter la lecture, on crée le tableau contenant les dossiers n'ayant pas été traités depuis + d'1 mois
		df_1_mois = df.loc[df.date < self.current_date - datetime.timedelta(30)]
		df_1_mois.to_csv("YOUR_PATH\\dossiers_plus_1_mois.csv", sep=";", index=False, encoding="utf-8-sig")

		#Ainsi que celui des dossiers n'ayant pas été traités depuis + de 3 mois
		df_3_mois = df.loc[df.date < self.current_date - datetime.timedelta(90)]
		df_3_mois.to_csv("YOUR_PATH\\dossiers_plus_3_mois.csv", sep=";", index=False, encoding="utf-8-sig")

		self.state_label.config(text="PROGRAMME EXECUTE AVEC SUCCES", fg="green")
		self.app.update()

	#Delete the bot and return to the cmd
	def destroy(self):
		self.driver.quit()
		os._exit(0)


#Fonction compte à rebours avant exécution automatique du programme de récupération de mails
def countdown(duration=60):
	t = duration
	start = time.time()
	while t > 0 and not stop_automatic[0]:
		t_temp = math.ceil(duration - (time.time() - start))
		if t != t_temp:
			t = t_temp
			remaining_time_label.config(text="Temps restant : {} secondes".format(t))
			app.update()

	if not stop_automatic[0]:
		stop_automatic[0] = True
		stop_manual[0] = True
		remaining_time_label.config(text="")
		state_label.config(text="EXECUTION AUTOMATIQUE CHOISIE")
		app.update()
		bot.get_sent_emails()

#Fonction d'exécution manuelle du programme (en appuyant sur valider)
def manual_execution():
	nb_days = nb_days_text.get()
	if nb_days.isnumeric():
		nb_days = int(nb_days)
		if not stop_automatic[0]:
			stop_automatic[0] = True
			if not stop_manual[0]:
				remaining_time_label.config(text="")
				state_label.config(text="EXECUTION MANUELLE CHOISIE")
				app.update()
				bot.get_sent_emails(nb_days)

#Partie Graphique
#Création de la fenêtre
app = tk.Tk()
fontStyle = tkFont.Font(size=14)

#Section "Nombre de jours"
nb_days_text = tk.StringVar()
nb_days_label = tk.Label(app, text="Jusqu'à combien de jours souhaitez-vous remonter ?", font=fontStyle)
nb_days_label.grid(row=0, column=0, sticky=W)
nb_days_entry = tk.Entry(app, textvariable=nb_days_text)
nb_days_entry.grid(row=0, column=1)

#Section "Temps restant"
remaining_time_label = tk.Label(app, font=(10))
remaining_time_label.grid(row=1, column=0, sticky=W)

#Bouton "Valider" à côté de la section "Nombre de jours"
exec_btn = tk.Button(app, text='Valider', width=14, command=manual_execution)
exec_btn.grid(row=0, column=2)

#Section "Déroulement"
deroulement_label = tk.Label(app, text="Déroulement", font=fontStyle)
deroulement_label.grid(row=2, column=0, pady=(30, 5), sticky=W)
dossier_list = tk.Listbox(app, height=20, width=150, border=2)
dossier_list.grid(row=3, column=0, columnspan=3, sticky=W)
#Création de la scrollbar et fixation à la listbox
scrollbar_dossier = tk.Scrollbar(app, orient=VERTICAL, command=dossier_list.yview)
dossier_list.configure(yscrollcommand=scrollbar_dossier.set)
scrollbar_dossier.grid(row=3, column=3, sticky=NS)

#Section "Avertissements"
warning_label = tk.Label(app, text="Avertissements", font=fontStyle)
warning_label.grid(row=4, column=0, pady=(30, 5), sticky=W)
warning_list = tk.Listbox(app, height=10, width=150, border=2)
warning_list.grid(row=5, column=0, columnspan=3, sticky=W)
#Création de la scrollbar et fixation à la listbox
scrollbar_warning = tk.Scrollbar(app, orient=VERTICAL, command=warning_list.yview)
warning_list.configure(yscrollcommand=scrollbar_warning.set)
scrollbar_warning.grid(row=5, column=3, sticky=NS)

#Section "Erreurs"
error_label = tk.Label(app, text="Erreurs", font=fontStyle)
error_label.grid(row=6, column=0, pady=(30, 5), sticky=W)
error_list = tk.Listbox(app, height=5, width=150, border=2)
error_list.grid(row=7, column=0, columnspan=3, sticky=W)
#Création de la scrollbar et fixation à la listbox
scrollbar_error = tk.Scrollbar(app, orient=VERTICAL, command=error_list.yview)
error_list.configure(yscrollcommand=scrollbar_error.set)
scrollbar_error.grid(row=7, column=3, sticky=NS)

#Section "ETAT"
new_fontStyle = tkFont.Font(weight="bold", size=14)
state_label = tk.Label(app, text="CHOIX DU NOMBRE DE JOURS...", font=new_fontStyle, fg="blue")
state_label.grid(row=8, column=0, pady=30, sticky=W)

app.title('Client File Manager')
app.geometry('952x981+942+10')


#PROGRAM

bot = BotCabinet(app, remaining_time_label, dossier_list, warning_list, error_list, state_label)
bot.login()
#On crée des listes, et non des booléens, afin de pouvoir les modifier à l'intérieur des fonctions
stop_automatic, stop_manual = [False], [False]
countdown()
app.mainloop()
#bot.destroy()
