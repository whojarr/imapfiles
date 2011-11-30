#!/usr/bin/python
"""
Copyright (C) 2011 by Digital Creation Ltd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.


This script will connect to an imap server and download the attachements from a specified folder
then archive them to a specified folder.
"""

import email
import ConfigParser
import imaplib
import datetime
import os
import sys
import optparse

def connect(server, username, password):
    """ connecting to the imap server """
    connection = imaplib.IMAP4_SSL(server)
    connection.login(username, password)
    return connection

class ImapFiles:
    """ Main class that connects to imap server and downloads attachments """
    def __init__(self):
        self.destination_folder = '.'
        self.config_file = None
        self.server = 'imap.gmail.com'
        self.username = ''
        self.password = ''
        self.imap_folder = ''
        self.imap_folder_archive = None

    def load_config_file(self, config_file):
        """ Set values used from a specified config file """
        if not os.path.isfile(config_file):
            sys.stderr.write('ERROR: file %s not found\n' % config_file)
            return 0
        else:
            config = ConfigParser.RawConfigParser()
            config.read(config_file)
            try:
                self.destination_folder = config.get('general','destination')
                self.server = config.get('imap', 'server')
                self.username = config.get('imap', 'username')
                self.password = config.get('imap', 'password')
                self.imap_folder = config.get('imap', 'folder')
                self.imap_folder_archive = config.get('imap', 'archive')
            except ConfigParser.NoOptionError, error:
                print "Config File Error: {0}".format(error)
                return 0

    def download(self):
        """ Download attachments from the imap account """
        if self.config_file != None:
            if not self.load_config_file(self.config_file):
                return 0
        connection = connect(self.server, self.username, self.password)
        connection.select(mailbox=self.imap_folder)
        response, items = connection.search(None, "ALL") #you could filter using the IMAP rules here (check http://www.example-code.com/csharp/imap-search-critera.asp)

        items = items[0].split() # getting the mails id
        for emailid in items:

            response, data = connection.fetch(emailid, "(RFC822)") # fetching the mail, "`(RFC822)`" means "get the whole stuff", but you can ask for headers only, etc
            #print('fetching email ' + emailid + " - " + response)

            email_body = data[0][1] # getting the mail content
            mail = email.message_from_string(email_body) # parsing the mail content to get a mail object

            #Check if any attachments at all
            if mail.get_content_maintype() != 'multipart':
                continue

            print "Email "+emailid+" FROM:"+mail["From"]+" SUBJECT:" + mail["Subject"]

            # we use walk to create a generator so we can iterate on the parts and forget about the recursive headach
            for part in mail.walk():
                # multipart are just containers, so we skip them
                if part.get_content_maintype() == 'multipart':
                    continue

                # is this part an attachment ?
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                counter = 1

                # if there is no filename, we create one with a counter to avoid duplicates
                if not filename:
                    filename = 'part-%03d%s' % (counter, 'bin')
                    counter += 1

                message_path = os.path.join(self.destination_folder, mail["From"], datetime.date.today().isoformat(), mail["message-id"])
                att_path = os.path.join(message_path, filename)

                print 'saving ' + att_path

                if not os.path.exists(message_path):
                    os.makedirs(message_path)

                #Check if its already there
                if not os.path.isfile(att_path) :
                    # finally write the stuff
                    newfile = open(att_path, 'wb')
                    newfile.write(part.get_payload(decode=True))
                    newfile.close()

            if(self.imap_folder_archive != None):
                response = connection.copy(emailid, self.imap_folder_archive) # move the message to an archive folder
                if response[0] == 'OK':
                    response = connection.store(emailid , '+FLAGS', '(\\Deleted)')

        connection.expunge()
        connection.logout()

if __name__ == '__main__':
    PARSER = optparse.OptionParser(usage = "Usage: %prog [OPTIONS]", version="%prog 0.1")
    PARSER.add_option("-s", "--server", dest="server", help="imap server name")
    PARSER.add_option("-c", "--config", dest="config", help="config file containing username, password, folder")
    PARSER.add_option("-u", "--username", dest="username", help="username for imap account")
    PARSER.add_option("-p", "--password", dest="password", help="passsword for imap account")
    PARSER.add_option("-f", "--folder", dest="folder", help="imap folder to search")
    PARSER.add_option("-a", "--archive", dest="archive", help="imap folder to archive emails that attachements downloaded from")
    PARSER.add_option("-d", "--destination", dest="destination", help="local folder to store downloaded attachments")

    OPTIONS, ARGS = PARSER.parse_args()
    
    IMAPFILES = ImapFiles()
    if not OPTIONS.config == None:
        IMAPFILES.load_config_file(OPTIONS.config)
        IMAPFILES.download()
    else:
        if not OPTIONS.server and not OPTIONS.username and not OPTIONS.password and not OPTIONS.folder and not OPTIONS.destination:
            PARSER.print_help()
            print('  Require -s SERVER -u USERNAME -p PASSWORD -f FOLDER -d DESTIATION if -c CONFIG not supplied')
            sys.exit(1)
        else:
            IMAPFILES.server = OPTIONS.server
            IMAPFILES.username = OPTIONS.username
            IMAPFILES.password = OPTIONS.password
            IMAPFILES.imap_folder = OPTIONS.folder
            if(OPTIONS.archive != None):
                IMAPFILES.imap_folder_archive = OPTIONS.archive
            IMAPFILES.destination_folder = OPTIONS.destination
            IMAPFILES.download()
