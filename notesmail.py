## A simple hack to send Lotus Notes email from Python

def sendEmail(recipients=[], subject='', body='', attachments=[]):
	"""Use Notes to send an email from the current user
	
	recipients -- a list of email addresses to send to 
		(or full names from the notes address book)
	subject -- a string containing the subject of the email
	body -- a string containing the body text of the email 
		(empty lines didn't seem to come through properly for me, I had to 
		include at least a space on each line to keep them from disappearing.)
	attachments -- a list of full path and filenames to attach to the email"""
	
	import win32com.client
	sess=win32com.client.Dispatch("Notes.NotesSession")
	db = sess.getdatabase('','')
	db.openmail
	doc=db.createdocument
	
	#Set the recipient to the current user as a default
	if not recipients:
		recipients = sess.UserName  
		
	doc.SendTo = recipients
	doc.Subject = subject
	doc.Body = body
	
	#Notes attachments get made in RichText items...
	if attachments:
		rt = doc.createrichtextitem('Attachment')
		for file in attachments:
			rt.embedobject(1454,'',file)
	doc.Send(0)
	
	
if __name__ == '__main__':
	#Send a simple test email to the current Notes user
	#sendEmail()
	sendEmail(subject='Test email from Python!', body='This has been a test, and only a test.')		
	