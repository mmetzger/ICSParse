#!/usr/bin/env python
#
# Copyright 2007 Google Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
#

import vobject
import StringIO
import datetime
import traceback
import re

import logging, email

from google.appengine.ext import webapp
from google.appengine.ext.webapp import util
from google.appengine.ext.webapp.mail_handlers import InboundMailHandler
from google.appengine.api import mail
from google.appengine.ext import db
from google.appengine.api import users

class MainHandler(webapp.RequestHandler):

  def get(self):
	self.response.out.write("""
		<html>
		  <head>
		    <title>Calendar Attachment Parser</title>
		  </head>
		  <body>
		    <h1>Calendar Attachment Parser</h1>
		    <p>
			  <ul>
				<li>What is it?
				  <p>Ever open an email on your iPhone and only see a calendar icon, but when you click on it you can't read anything?  This site is designed to mitigate that problem until Apple decides to fix ICS file parsing.</p>
				</li>
				<li>What do I do?
				  <p>It's pretty simple - just forward any message containing an ics file to <a href="mailto:beta@icsparse.appspotmail.com">beta@icsparse.appspotmail.com</a>.  The service will parse your message and will e-mail you back the details in a fairly human-readable format.  Make sure to include the original attachment in the forwarded mail - the iPhone sometimes takes longer to download this and send it.  You can verify the file is attached by viewing the icon at the bottom of the email message.</p>
				</li>
				<li>Awesome!  How much?
				  <p>It's free!  Please feel free to use the service as needed, but try not to overload it.</p>
				</li>
				<li>The times seem off?
				  <p>This is an issue with the way a lot of clients handle the timezone representation within the ICS files.  The iCalendar spec allows for timezone / offsets to be put in the date but some clients do not do this (Outlook especially.)  Usually, the easiest thing is to look in the Description section for the actual time / date of the meeting, otherwise we present it as listed in the file.  Typically, this is set for GMT.  We have some ideas on handling it in the future, but for now just be careful.</p>
				</li>
				<li>How do I add it to my calendar?
				  <p>We're working on that for devices running the 4.0+ iOS (iPhone / iPod Touch)!  Still some kinks to work out but we hope to have a release soon.</p>
				</li>
				<li>My e-mail was rejected...
				  <p>Make sure you sent it to <a href="mailto:beta@icsparse.appspotmail.com">beta@icsparse.appspotmail.com</a> - currently any other address is set to automatically reject email requests.</p>
				</li>
				<li>What about an iPhone app?
				  <p>We've got one for users with iOS4 or later!  ICS Reader is now available on the <a href="http://itunes.apple.com/us/app/ics-reader/id396384094?mt=8">App Store</a> for free!  More information also available at <a href="http://icsreader.com">http://icsreader.com</a>.</p>
				</li>
				<li>Cool! Does that mean I can open every ics file automatically?
				  <p>Yes and no - if you're using something like Dropbox or another app that supports the a document selection viewer, then yes, it will work well.  If you're wanting to open it from e-mail (95% of you out there) then there's still a minor workaround required but it does work.</p>
				</li>
				<li>What if I don't have an iPhone?
				  <p>In most cases, you shouldn't need this as the mail clients can parse it (Note: the iPhone Mail.app knows how to handle ICS files - the problem is it just doesn't try when you're connected to an IMAP or POP3 service, only Exchange or MobileMe.)  The service should work with any given e-mail client so if you've got something that can't parse the file but can read text or html, it should work for you.</p>
				</li>
				<li>How can I accept / decline / etc the calendar invitation?
				  <p>We're looking into this - it's definitely possible to do, it's mostly a matter of sending back the proper type of ICS file in response which can be a bit tricky.  Please let us know if this is something you need.</p>
				</li>
				<li>What about privacy / reliability?
				  <p>We have no intention of using your data unless there's an error / problem / etc.  That said, we cannot and do not promise the security of any information sent to this service.  If you're concerned in any way about this, DO NOT USE THIS SERVICE.  In addition, we make no guarantees as to the availability of the service nor it's effectiveness.  It works fine for us (make sure to read about the timezone piece above) but it may not work for you.  When in doubt, just open your laptop and verify the settings that way.  Again, if any of this concerns you, DO NOT USE THIS SERVICE.</p>
				</li>
				<li>What is this site / service built on?
				  <p>This service is built using python running on top of Google AppEngine.  We specifically use the great <a href="http://vobject.skyhouseconsulting.com/">VObject</a> library to handle parsing of the ics files.  Everything else is built into AppEngine, mostly just code to attempt to intelligently handle the mutilation of attachments received from various e-mail clients.  For the coming iOS app, we've had to change things somewhat - details to follow.</p>
				</li>
				<li>Something doesn't work! Help!
				  <p>First, check to make sure the ics file is actually attached - if the service does not detect an attachment it will not send any response.  If you sent a file with an attachment, you should receive a response even if it states it couldn't parse the attachment.  In these cases, if you're comfortable doing so, please forward us a copy of the e-mail with the attachment and we'll look into it.  Our e-mail address is <a href="mailto:icsparse@flexiblecreations.com">icsparse@flexiblecreations.com</a>.</p><p>Note that we may also have the site in various states of development at any given time so the service may not be active.</p><p>It's also come to our attention that in certain circumstances, the iPhone likes to hold onto the e-mail address that replies to you, meaning icsparse@flexiblecreations.com.  This is a side-effect of using Google AppEngine and we have a fix in development.  In the mean time, just make sure you're using the beta@icsparse.appspotmail.com address.  Don't worry if it happens accidentally, we don't mind - we just want to make sure the service is working well for you and we can't always reply immediately if it gets sent to the wrong address.</p>
				</li>
				<li>I've got a question about something not listed here...
				  <p>We welcome feedback and questions about the service - this is largely an experiment and to solve a couple pain points we experience.  Feel free to contact us at the address listed above.</p>
				</li>
			  </ul>
		    </body>
		</html>""")

class ReceivedMessage(db.Model):
	to = db.StringProperty(multiline=True)
	subject = db.StringProperty(multiline=True)
	received = db.DateTimeProperty()
	body = db.StringProperty()
	html = db.StringProperty()
	account = db.StringProperty()
	attachment = db.BlobProperty()

			
class MailHandler(InboundMailHandler):
  def receive(self, message):
	logging.info("Recevied a message from: " + message.sender)
	r = ReceivedMessage()
	
	logging.info('To: %r', message.to)
	r.to = message.to
	if hasattr(message, 'subject'):
		r.subject = message.subject
	else:
		r.subject = "Response from ICS Parser"
	r.received = datetime.datetime.now()
	#r.body = message.body
	#r.html = message.html
	r.account = message.sender
	
	attachments = []
	try:
		if message.attachments is not None:
			if isinstance(message.attachments[0], basestring):
				attachments = [message.attachments]
			else:
				attachments = message.attachments
		logging.info(attachments)
		body = u"\n"
		reply_to = u"beta@icsparse.appspotmail.com"
		count = 1
		calendarattachments = []
		for fn, content in attachments:
			body += u"Attachment " + unicode(count) + u":\n"
			body += u"Filename: " + unicode(fn) + u"\n"
			try:
				payload = goodDecode(content)
				#			r.attachment = db.Blob(payload)

				r.put()
				ics = vobject.readOne(StringIO.StringIO(payload))
				# Start pulling out other info
				# Need to handle timestamp issues / etc
				# Send data
				# Send ads with event data?
				#
				if hasattr(ics.vevent, 'organizer'):
					body += u"Organizer: <a href=\"" + unicode(ics.vevent.organizer.value) + u"\">"
					reply_to = unicode(ics.vevent.organizer.value).replace(u'mailto:', u'')
					if u'CN' in ics.vevent.organizer.params:
						body += unicode(ics.vevent.organizer.params[u'CN'][0])
				 	else:
						body += unicode(ics.vevent.organizer.value)
					body += u"</a>\n"
				if hasattr(ics.vevent, 'summary'):
					body += u"Summary: " + unicode(ics.vevent.summary.value) + u"\n"
				if hasattr(ics.vevent, 'dtstart'):
					body += u"Starts: " + unicode(ics.vevent.dtstart.value.strftime("%A, %d. %B %Y %I:%M%p")) + u"\n"
				if hasattr(ics.vevent, 'dtend'):
					body += u"Ends: " + unicode(ics.vevent.dtend.value.strftime("%A, %d. %B %Y %I:%M%p")) + u"\n"
				if hasattr(ics.vevent, 'location'):
					body += u"Location: " + unicode(ics.vevent.location.value) + u"\n"
				if hasattr(ics.vevent, 'description'):
					body += u"Description:\n"
					body += unicode(ics.vevent.description.value)
				body += u"\n\n"
				calendarattachments.append((fn, content))
			except Exception, e:
				logging.debug(traceback.format_exc(e))
				body += u"Could not parse this attachment\n\n"
			count += 1

		if count > 1:
					body += u"\n\n--------------------------\n\nIt's out!  You can now download the ICS Parser on the App Store.  Read ICS files and add them to your calendar with a click on the link.  Make sure to read the instructions on how it works, especially if you're reading the attachments in Mail.  The best part - it's FREE!  Get it from the App Store at http://itunes.apple.com/us/app/ics-reader/id396384094?mt=8.  More info at icsreader.com.\n\n"
					#returnattachments = [('holderfile.pdf', '0000')]
					#returnattachments.extend(calendarattachments)
					htmlbody = re.sub('\n', '<br>\n', body)
					mail.send_mail(sender="ICS Parser <icsparse@flexiblecreations.com>", to=message.sender, reply_to = reply_to, subject="Calendar parsing", body=body, html=htmlbody) #, attachments=returnattachments)
	except Exception, e:
		logging.debug(traceback.format_exc(e))

def goodDecode(encodedPayload):
	if not hasattr(encodedPayload, 'encoding'):
		return encodedPayload
	else:
		encoding = encodedPayload.encoding
		payload = encodedPayload.payload
		if encoding and encoding.lower() != '7bit' and encoding.lower() != '8bit':
			payload = payload.decode(encoding)
		return payload


def main():
  application = webapp.WSGIApplication([('/', MainHandler), 
										MailHandler.mapping()],
										debug=True)
  util.run_wsgi_app(application)


if __name__ == '__main__':
  main()
