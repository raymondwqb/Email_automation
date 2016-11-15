import smtplib
import imaplib
import poplib
import getpass
import email
import email.utils
import socket
import sys
import os
import json
import collections
import datetime
import numpy

class emails(object):
	"""docstring for emails"""
	def __init__(self):
		super(emails, self).__init__()
		self.email_address = self.prompt("Email Address:")
		self.password = self.getpass("PassWord:")
		self.imap_host = 'imap.qq.com'
		self.imap_port = 993
		self.box = 'inbox'
		self.smtp_host = 'smtp.qq.com'
		self.smtp_port = 465

	def prompt(self, sentence):
		return raw_input(sentence).strip()

	def getpass(self, sentence):
		return getpass.getpass(sentence)

	def imap_connect(self):
		if self.imap_port:
			conn = imaplib.IMAP4_SSL(self.imap_host, self.imap_port)
		else:
			conn = imaplib.IMAP4_SSL(self.imap_host)
		conn.login(self.email_address, self.password)
		return conn

	def imap_reconnect(self):
		print("Current connection Lost... Trying to reconnect ...")
		conn = self.imap_connect()
		conn.select(self.box)
		return conn

	def imap_disconnect(self, conn):
		conn.close()
		conn.logout()
		return 0

	def smtp_disconnect(self, conn):
		conn.quit()
		return 0

	def smtp_connect(self):
		if self.smtp_port:
			conn = smtplib.SMTP_SSL(self.smtp_host, self.smtp_port)
		else:
			conn = smtplib.SMTP_SSL(self.smtp_host)
		conn.login(self.email_address, self.password)
		return conn

	def smtp_reconnect(self):
		print("Current connection Lost... Trying to reconnect ...")
		conn = self.smtp_connect()
		return conn

	def fetch_senders(self, conn):
		sender_list = []
		result, data = conn.search(None, 'ALL')
		if result != 'OK':
			raise ValueError("Can Not Get Mail List")
		else:
			try:
				result, mails = conn.uid('fetch', ','.join(data[0].split()), "(RFC822)")
			except:
				conn = self.imap_reconnect()
				result, mails = conn.uid('fetch', ','.join(data[0].split()), "(RFC822)")
			for body in mails:
				headers = email.parser.HeaderParser().parsestr(body[0]) 
				address = headers['From']
				sender_list.append(email.utils.parseaddr(address)[-1])
				result_list.append(result)
		return ('OK', list(set(sender_list))) 	

	def fetch_sender(self, conn, id):
		try:
			result, body = conn.fetch(id, "(RFC822)")
		except:
			conn = self.imap_reconnect()
			result, body = conn.fetch(id, "(RFC822)")
			headers = email.parser.HeaderParser().parsestr(body[0][1]) 
			address = headers['From']
		else:
			headers = email.parser.HeaderParser().parsestr(body[0][1]) 
			address = headers['From']

		return result, email.utils.parseaddr(address)[-1]

	def get_latest_id(self, conn):
		try:
			result, data = conn.search(None, 'ALL')
			if result != 'OK':
				raise ValueError("Can Not Get Mail List")
		except:
			conn = self.imap_reconnect()
			result, data = conn.search(None, 'ALL')
		return data[0].split()[-1]

	def get_email(self, conn, id):
		try:
			status, data = conn.fetch(id, "(RFC822)")
			if status != 'OK':
				raise ValueError("Can Not Get Mail List")
		except:
			conn = self.imap_reconnect()
			status, mail = conn.fetch(id, "(RFC822)")
			message = email.message_from_string(mail[0][-1])
		else:
			message = email.message_from_string(mail[0][-1])
		return message

	def fetch_all_senders(self, conn, num = None):
		sender_list = []
		status_list = []
		try:
			status, data = conn.search(None, 'ALL')
			if status != 'OK':
				raise ValueError("Can Not Get Mail List")
		except:
			conn = self.imap_reconnect()
			status, data = conn.search(None, 'ALL')
		lst = data[0].split()
		if num:
			if isinstance(num, (int, long)):
				pass
			elif isinstance(num, (str)):
				if num.isdigit():
					num = int(num)
				else:
					num = len(lst) - 1
			else:
				num = len(lst) - 1
		else:
			num = len(lst) - 1
		num = min(num, len(lst) - 1)
		for id in lst[:num]:
			try:
				status, address = self.fetch_sender(conn, id)
			except:
				conn = self.imap_reconnect()
				status, address = self.fetch_sender(conn, id)

			if address not in sender_list:
				sender_list.append(address)
				status_list.append(status)

		if not all(status == 'OK' for status in status_list):
			raise ValueError("Can Not Get Mail List")
		return 'OK', sender_list

	def get_latest_email(self, conn):
		try:
			id = self.get_latest_id(conn)
		except:
			conn = self.imap_reconnect()
			id = self.get_latest_id(conn)
		try:
			status, mail = conn.fetch(id, "(RFC822)")
		except:
			conn = self.imap_reconnect()
			status, mail = conn.fetch(id, "(RFC822)")
		finally:
			message = email.message_from_string(mail[0][-1])
		return message

	def forward_email(self, conn, msg, email_from, email_to):
		msg.replace_header("From", email_from)
		for email_add in email_to:
			msg.replace_header("To", email_add)
			try:
				conn.sendmail(email_from, email_add, msg.as_string())
				print "Email Sent to %s" %(email_add)
			except:
				conn = smtp_reconnect()			
				conn.sendmail(email_from, email_add, msg.as_string())
				print "Email Sent to %s" %(email_add)
		conn.quit()
		return 0

	@staticmethod
	def save_list(lst, filename):
		if not os.path.exists(os.path.dirname(os.path.abspath(filename))):
			os.makedirs(os.path.dirname(os.path.abspath(filename)))		
		files = open(os.path.abspath(filename), 'w')
		file.write("\n".join(lst))
		file.close()
		return 0

	@staticmethod
	def read_list(filename):
		if not os.path.exists(os.path.abspath(filename)):
			raise ValueError("File Doesn't exist %s" %(filename))
		lst = []
		files = open(os.path.abspath(filename), 'r')
		for line in files:
			lst.append(line)
		return lst

if __name__ == "__main__":
	socket.setdefaulttimeout(100)
	print("Starting Connecting to IMAP Server ...")
	conf = emails()
	imap_connection = conf.imap_connect()
	imap_connection.select(conf.box)
	print("Starting Fetching all email senders from IMAP Server ...")
	senders = conf.fetch_all_senders(imap_connection, 100)
	filename = 'senders.txt'
	print("Starting Saving all email senders list to file %s..." %(os.path.abspath(filename)))
	emails.save_list(senders, filename)
	print("Starting Fetching the latest email body from IMAP Server ...")
	message = conf.get_latest_email(imap_connection)
	print("Starting Connecting to SMTP Server ...")
	smtp_connection = smtp_connect()
	print("Starting Forwarding the latest email to mailing list from SMTP Server ...")
	status = conf.forward_email(smtp_connection, message, conf.email_address, senders)
	if status != 0:
		logger.info("Failure")
		raise ValueError("Failed to send emails")
	else:
		logger.info("Success")

	