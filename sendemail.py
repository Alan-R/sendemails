#!/usr/bin/env python3
# vim: smartindent tabstop=4 shiftwidth=4 expandtab number colorcolumn=80
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
and seem a bit more homey and human-created.

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
                destinationcsv - name of CSV file as noted below

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
import re, smtplib, os, time, sys, pytz
from datetime import datetime
from email.mime.text import MIMEText

allnames = {}
allemails = {}

def format_and_send_email(text, subject, smtpinfo, keys, flags=None):
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
    subject = subject.strip()
    flags = flags or {}
    dontsend = flags.get('dontsend', False)
    keywords = keys
    dest = keywords['email']
    name = keywords['name']
    if dest.lower() in allemails:
        print ('OOPS: Email address %s is duplicate' % dest)
        return
    allemails[dest.lower()] = name.lower()
    if name.lower() in allnames:
        print ('OOPS: Name %s is duplicate' % name)
        return
    allnames[name.lower()] = dest.lower()
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
    if dest.find('@') == -1:
        raise ValueError('Email address for %s ["%s"] is invalid' %
                         (name, dest))
    tz = pytz.timezone(keywords['timezone'])
    keywords['Date'] = datetime.now(tz).strftime('%F %T %z')
    if 'Today' not in keywords:
        keywords['Today'] = datetime.now(tz).strftime('%d %B')
    if 'Year' not in keywords:
        keywords['Year'] = datetime.now(tz).strftime('%Y')
    if dontsend:
        outtext = substitute_text(text, keywords)
        outsubject = substitute_text(subject, keywords)
        try:
            pytz.timezone(keywords['timezone'])
            #print('    %s OK.' % (toaddr))
        except pytz.exceptions.UnknownTimeZoneError:
            print('ERROR: Time Zone for %s [%s] is invalid'
                  % (toaddr, keywords['timezone']))

    if should_send_now(keywords, flags):
        outtext = substitute_text(text, keywords)
        outsubject = substitute_text(subject, keywords)
        send_an_email(toaddr, outsubject, smtpinfo, outtext, keywords)

WEEKDAYS = ('monday', 'tuesday', 'wednesday', 'thursday',
            'friday', 'saturday', 'sunday')

def should_send_now(keywords, flags):
    '''
    Return TRUE if we should send this message now...

    We are given a time zone ['timezone'] and an hour ['sendhour']
    to send this message, and if it's the requested hour in that time zone,
    then it's time to send the message.

    If we're not given a 'sendhour', it's always time to send the message ;-).

    '''
    flags = flags or {}
    debug = flags.get('debug', False)
    if 'sendhour' not in keywords:
        return True
    ourtz = pytz.timezone(keywords['timezone'])
    hour = datetime.now(ourtz).hour
    if debug:
        print('In %s, today is %s, hour is %s.'
              % (keywords['timezone'],
                 WEEKDAYS[datetime.now(ourtz).date().weekday()].title(),
                 int(hour)))
    if hour != int(keywords['sendhour']):
        return False
    if 'sendday' not in keywords:
        return True
    sendday = keywords['sendday']
    intday = None
    for i in range(0, len(WEEKDAYS)):
        if sendday.lower() == WEEKDAYS[i]:
            intday = i
            break
    if intday is None:
        raise ValueError("Illegal sendday: %s" % sendday)
    return intday == datetime.now(ourtz).date().weekday()

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

def send_an_email(toaddr, subject, smtpinfo, msgbody, keywords, smtpdebug=False):
    '''
    We need this info in the smtpinfo:
        gateway     System providing SMTP service
        login       Login name for 'gateway'
        password    password for 'login'
        from        From address for message header

	Here are some things we'd like to find in smptinfo
Organization: Assimilation Systems Limited
Message-ID: <55F32FF9.9000003@unix.sh>
Date: Fri, 11 Sep 2015 13:48:09 -0600
User-Agent: Mozilla/5.0 (X11; Linux x86_64; rv:38.0) Gecko/20100101 Thunderbird/38.2.0

    '''
    addrmatch = DESTPATTERN.search(toaddr)
    if addrmatch:
        dest = addrmatch.group()
    else:
        dest = toaddr
    msg = MIMEText(msgbody, 'plain', 'utf-8')
    msg['To'] = toaddr
    msg['From'] = smtpinfo['from']
    msg['Subject'] = subject
    for header in ('Organization', 'Message-ID', 'Date', 'User-Agent'
                   'Importance'):
        if header in smtpinfo:
            msg[header] = smtpinfo[header]

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
    with open(filename, mode='r', encoding='utf-8') as smtpfile:
        count = 0
        while True:
            line = smtpfile.readline().strip()
            count += 1
            if line == '':
                break
            if line.startswith('#'):
                continue
            try:
                name, value = line.split('=', 1)
            except ValueError:
                raise ValueError('SMTP line %d "%s" is incorrect'
                                 % (count, line.strip()))

            keywords[name] = value
    return keywords


def process_csv_file(csvfilename, action, smtpkw, flags=None):
    '''
    Apply the given action to each line of the CSV file we've been given.
    '''
    with open(csvfilename, mode='r', encoding='utf-8') as csvfile:
        initline = csvfile.readline()
        keywords = initline.strip().split(',')
        while True:
            csvkw = {}
            line = csvfile.readline().strip()
            if line == '':
                return
            if line.startswith('#'):
                continue
            linewords = line.strip().split(',')
            if len(linewords) != len(keywords):
                raise ValueError("Line %s has %s elements instead of %s"
                                 % (line, len(linewords), len(keywords)))
            for j in range(0, len(keywords)):
                csvkw[keywords[j]] = linewords[j]
            action(csvkw, smtpkw, flags=flags)

def send_emails_to_csv_people(ourkw, smtpkw, flags=None):
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
            raise(ValueError("Message in file %s is too old (%s hours)."
                             % (bodyfile, ageinhours)))
    with open(bodyfile, mode='r', encoding='utf-8') as plainbody:
        subject = plainbody.readline()
        plaintext = plainbody.read()
        format_and_send_email(plaintext, subject, smtpkw, ourkw, flags)

def test_emails_to_csv_people(ourkw, smtpkw, flags):
    '''
    Test Action function for 'process_csv_file' to validate
    our CSV file and message.

    ourkw is the keywords for this particular person
    smtpkw is the set of (SMTP) keywords that are common to all emails.

    It might be nice to validate each email address - here's a
    suggestion on how to do it:
        https://gist.github.com/blinks/47987
    based on the technique described here:
    https://www.webdigi.co.uk/blog/2009/how-to-check-if-an-email-address-exists-without-sending-an-email/

    The Webdigi page mentions that if you do this continually to check
    for gmail/yahoo/msn accounts this may cause your IP to be added
    to a blacklist. Not a good thing!

    Not only that I suspect this technique won't work for some of these
    sites because they likely only reject the email after it's been sent.
    There are currently 72 gmail addresses in the list. That's a quite a few.

    Tried the technique by hand, could not make it work for any emails...
    I wonder if my home address is in a block list?
    '''
    bodyfile = smtpkw['plainbody']
    with open(bodyfile, mode='r', encoding='utf-8') as plainbody:
        subject = plainbody.readline()
        plaintext = plainbody.read()
        ourflags = flags
        ourflags['dontsend'] = True
        format_and_send_email(plaintext, subject, smtpkw, ourkw, ourflags)

def maintest():
    'Main test program'
    smtpfile = 'smtp.txt'
    legalflags = {'debug', 'test'}
    flags = {}
    for flag in legalflags:
        flags[flag] = False

    for arg in sys.argv[1:]:
        if arg.startswith('--'):
            flag = arg[2:]
            if flag not in legalflags:
                raise ValueError("Illegal flag '%s': legal flags are %s" %
                                 (flag, str(legalflags)))
            flags[flag] = True
        else:
            smtpfile = arg
    smtpkw = get_smtpinfo(smtpfile)
    sendfunc = (test_emails_to_csv_people if flags['test']
                else send_emails_to_csv_people)
    process_csv_file(smtpkw['destinationcsv'], sendfunc, smtpkw, flags)

    if flags.get('test', False):
        if 'sendhour' in smtpkw:
            print('Sending on %s at %02dxx. Today is %s in local time zone.'
                  % (smtpkw['sendday'], int(smtpkw['sendhour']),
                     WEEKDAYS[datetime.now().date().weekday()].title()))
        print('Email tests on %s=>%s=>%s complete.'
              % (smtpfile, smtpkw['plainbody'], smtpkw['destinationcsv']))
        print('Failures (if any) noted above.')

if __name__ == '__main__':
    maintest()
