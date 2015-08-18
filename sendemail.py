#!/usr/bin/env python3
# vim: smartindent tabstop=4 shiftwidth=4 expandtab number
#
# This file is part of the Assimilation Project.
#
# Author: Alan Robertson <alanr@unix.sh>
# Copyright (C) 2015 - Alan Robertson
#
#
# This software is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# The Assimilation software is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with the Assimilation Project software.  If not, see
# http://www.gnu.org/licenses/
#
'''
Module (main program) for sending email to a collection of people
while performing substitution to make the emails personalized
and seem homey and human-created.

We use two different files for our input:
    SMTP information - formatted as name=value
            The following fields are required to be provided in this file:
                gateway   - SMTP system to send mail through
                login     - login name to use when connecting to gateway
                password  - password to use when connecting to gateway
                plainbody - name of file containing the plain ASCII text
                            to be sent as the email. It can contain keywords
                            to be substituted from the CSV or SMTP information
                            as @@keyword@@.
                from      - From address to put in outgoing message header

        Sample SMTP information (but don't indent)
            from=Jacob Marley <JacobMarley@ScroogeWorks.com>
            gateway=smtp.scroogeworks.com
            login=jacob@scroogeworks.com
            password=a Christmas Carol by Charles Dickens
            plainbody=christmas-email.txt
        Comments are lines starting with #

    CSV information naming recipients and recipient keywords as described below.

The key is to have a CSV file created with headers which define all the things
you want to substitute into the message. This follows the CSV convention
of the first line defining the column labels (headers). It might look something
like this:
    email,name,organization,timezone

The fields 'email' and 'name' are required. If you do not have a 'firstname'
field, then it will be created from the 'name' field.
You can have comment lines in this file (starting with #), but they CANNOT
appear before the first (header) line.
'''
KWDELIM = '@@'
import re, smtplib, os, time
from datetime import datetime
from pytz import timezone
from email.mime.text import MIMEText

def format_and_send_email(text, subject, smtpinfo, keys):
    '''
    Format and send an email based on the given text, subject and keywords.
    The following keywords have well-known meanings.
        email       The email address to send the email to (required)
        name        The full name (if you know it) of the person sending this
                    message to (required)
        firstname   The First name (if you know it) for the person you're
                    sending this message to (optional)

        Any @@keywords@@ which are given will be substituted everywhere they
        are found in the subject or body
        Keywords are defined by either the 'keys' or 'smtpinfo' arguments.
        This allows for common substitutions for everyone (in smptinfo) as well
        as substitutions on a per-destination-user basis.
    '''
    keywords = keys
    dest = keywords['email']
    name = keywords['name']
    if 'firstname' in keywords:
        firstname = keywords['firstname']
    else:
        firstname = name.split(' ', 1)[0]
        keywords['firstname'] = firstname
    if "'" in name:
        toaddr = '"%s" <%s>"' % (name, dest)
    else:
        toaddr = '%s <%s>' % (name, dest)
    for key in smtpinfo:
        if (key not in ('login', 'password', 'plainbody', 'htmlbody')
                and key not in keywords):
            keywords[key] = smtpinfo[key]
    if should_send_now(keywords):
        outtext = substitute_text(text, keywords)
        outsubject = substitute_text(subject, keywords)
        send_an_email(toaddr, outsubject, smtpinfo, outtext)

def should_send_now(keywords):
    '''
    Return TRUE if we should send this message now...

    We are given a time zone ['timezone'] and an hour ['sendhour']
    to send this message, and if it's the requested hour in that time zone,
    then it's time to send the message.

    If we're not given a 'sendhour', it's always time to send the message ;-).

    '''
    if 'sendhour' not in keywords:
        return True
    hour = datetime.now(timezone(keywords['timezone'])).hour
    return hour == int(keywords['sendhour'])


def substitute_text(text, keys):
    '''
    We substitute keywords where they may be found in the given text
    like this: @@keyword-name@@.
    Any keywords found in the text, but not found in the 'kw' parameter
    will cause an exception to be raised.
    '''
    outtext = text
    keywords_in_text = find_keywords(text)
    for keyword in keywords_in_text:
        if keyword not in keys:
            raise  ValueError("Undefined keyword '%s' found in email template."
                              % keyword)
        outtext = outtext.replace('%s%s%s' %
                                  (KWDELIM, keyword, KWDELIM), keys[keyword])
    return outtext

def find_keywords(text):
    '''
    Find keywords that look like @@keyword@@, and return the set of
    keywords we found. List or set would do, but we currently return a list...
    '''
    kwpat = re.compile('%s([^%s]+)%s' % (KWDELIM, KWDELIM, KWDELIM))
    return kwpat.findall(text)

DESTPATTERN = re.compile('<(.*)>')

def send_an_email(toaddr, subject, smtpinfo, msgbody, smtpdebug=False):
    '''
    We need this info in the smtpinfo:
        gateway     System providing SMTP service
        login       Login name for 'gateway'
        password    password for 'login'
        from        From address for message header
    '''
    addrmatch = DESTPATTERN.search(toaddr)
    if addrmatch:
        dest = addrmatch.group()
    else:
        dest = toaddr
    msg = MIMEText(msgbody)
    msg['To'] = toaddr
    msg['From'] = smtpinfo['from']
    msg['Subject'] = subject

    print("Sending email to %s." % toaddr)
    server = smtplib.SMTP(smtpinfo['gateway'], 587)
    try:
        server.set_debuglevel(smtpdebug)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(smtpinfo['login'], smtpinfo['password'])
        server.sendmail(dest, [toaddr], msg.as_string())
    finally:
        server.quit()

def get_smtpinfo(filename):
    '''
    We read our SMTP keywords from a file for sending info.
    This file should be mode 600 or 400.
    '''
    keywords = {}
    with open(filename, 'r') as smtpfile:
        while True:
            line = smtpfile.readline()
            if line == '':
                break
            if line.startswith('#'):
                continue
            name, value = line.strip().split('=', 1)
            keywords[name] = value
    return keywords


def process_csv_file(csvfilename, action, smtpkw):
    '''
    Apply the given action to each line of the CSV file we've been given.
    '''
    with open(csvfilename, 'r') as csvfile:
        initline = csvfile.readline()
        keywords = initline.strip().split(',')
        while True:
            csvkw = {}
            line = csvfile.readline()
            if line == '':
                break
            if line.startswith('#'):
                continue
            linewords = line.strip().split(',')
            if len(linewords) != len(keywords):
                raise ValueError("Line %s has %s elements instead of %s"
                                 % (line, len(linewords), len(keywords)))
            for j in range(0, len(keywords)):
                csvkw[keywords[j]] = linewords[j]
            action(csvkw, smtpkw)

def send_emails_to_csv_people(ourkw, smtpkw):
    '''
    Action function for 'process_csv_file'
    ourkw is the keywords for this particular person
    smtpkw is the set of (SMTP) keywords that are common to all emails.
    '''
    bodyfile = smtpkw['plainbody']
    if 'maxagehours' in smtpkw:
        fileage = time.time() - os.path.getmtime(bodyfile)
        ageinhours = fileage / (60*60)
        if ageinhours > int(smtpkw['maxagehours']):
            raise(ValueError("Message in file %s is too old to send out"
                             % bodyfile))
    with open(bodyfile, 'r') as plainbody:
        subject = plainbody.readline()
        plaintext = plainbody.read()
        format_and_send_email(plaintext, subject, smtpkw, ourkw)


def maintest():
    'Main test program'
    smtpkw = get_smtpinfo('smtp.txt')
    process_csv_file('destinations.csv', send_emails_to_csv_people, smtpkw)

if __name__ == '__main__':
    maintest()
