#!/usr/bin/env python3

"""
This script is meant as a simple way to reply to ical invitations from mutt.
See README for instructions and LICENSE for licensing information.
"""

__author__="Martin Sander"
__license__="MIT"

import vobject
import tempfile, time
import os, sys
import warnings
from datetime import datetime
from subprocess import Popen, PIPE
import argparse

parser = argparse.ArgumentParser(description="""
Parse an ics file and send a reply
""")

parser.add_argument('-e', metavar='email', help="Email address to send from")
intornot = parser.add_mutually_exclusive_group(required=True)
intornot.add_argument('-v', action='store_true', help="View the invite")
intornot.add_argument('-i', action='store_true', help="Interactive mode, ask for an answer")
answergroup = intornot.add_mutually_exclusive_group()
answergroup.add_argument('-a', action='store_true', default=True, help='Accept')
answergroup.add_argument('-d', action='store_true', help='Decline')
answergroup.add_argument('-t', action='store_true', help='Tentative')
parser.add_argument('icsfile', help="Path to the ics file")

def del_if_present(dic, key):
    if key in dic:
        del dic[key]

def set_accept_state(attendees, state):
    for attendee in attendees:
        attendee.params['PARTSTAT'] = [state]
        for i in ["RSVP","ROLE","X-NUM-GUESTS","CUTYPE"]:
            del_if_present(attendee.params,i)
    return attendees

def get_accept_decline():
    while True:
        ans = input("Accept Invitation? [y/n/t/Q]")
        if ans.lower() == 'q' or ans == '':
            sys.exit(1)
        elif ans.lower() == 'y':
            return 'ACCEPTED'
        elif ans.lower() == 'n':
            return 'DECLINED'
        elif ans.lower() =='t':
            return 'TENTATIVE'

def get_answer(invitation):
    # create
    ans = vobject.newFromBehavior('vcalendar')
    ans.add('method')
    ans.method.value = "REPLY"
    ans.add('vevent')

    # just copy from invitation
    #for i in ["uid", "summary", "dtstart", "dtend", "organizer"]:
    # There's a problem serializing TZ info in Python, temp fix
    for i in ["uid", "summary", "organizer"]:
        if i in invitation.vevent.contents:
            ans.vevent.add( invitation.vevent.contents[i][0] )

    # new timestamp
    ans.vevent.add('dtstamp')
    ans.vevent.dtstamp.value = datetime.utcnow().replace(
            tzinfo = invitation.vevent.dtstamp.value.tzinfo)
    return ans

def write_to_tempfile(ical):
    tempdir = tempfile.mkdtemp()
    icsfile = tempdir+"/event-reply.ics"
    with open(icsfile,"w") as f:
        f.write(ical.serialize())
    return icsfile, tempdir

def get_mutt_command(command_base, ical, email_address, accept_decline, icsfile):
    accept_decline = accept_decline.capitalize()
    if 'organizer' in ical.vevent.contents:
        if hasattr(ical.vevent.organizer,'EMAIL_param'):
            sender = ical.vevent.organizer.EMAIL_param
        else:
            sender = ical.vevent.organizer.value.split(':')[1] #workaround for MS
    else:
        sender = "NO SENDER"
    summary = ical.vevent.contents['summary'][0].value
    command = [ f % dict(icsfile=icsfile, summary=summary, sender=sender,
        accept_string=accept_decline) for f in
            command_base ]
    #"mutt", "-a", icsfile,
    #        "-s", "'%s: %s'" % (accept_decline, summary), "--", sender]
    #Uncomment the below line, and move it above the -s line to enable the wrapper
            #"-e", 'set sendmail=\'ical_reply_sendmail_wrapper.sh\'',
    return command

def execute(command, mailtext):
    process = Popen(command, stdin=PIPE)
    process.stdin.write(mailtext)
    process.stdin.close()

    result = None
    while result is None:
        result = process.poll()
        time.sleep(.1)
    if result != 0:
        input("unable to send reply, subprocess exited with\
                exit code %d\nPress return to continue" % result)

def openics(invitation_file):
    with open(invitation_file) as f:
        try:
            with warnings.catch_warnings(): #vobject uses deprecated Exception stuff
                warnings.simplefilter("ignore")
                invitation = vobject.readOne(f, ignoreUnreadable=True)
        except AttributeError:
            invitation = vobject.readOne(f, ignoreUnreadable=True)
    return invitation

def display(ical):
    summary = ical.vevent.contents['summary'][0].value
    if 'organizer' in ical.vevent.contents:
        if hasattr(ical.vevent.organizer,'EMAIL_param'):
            sender = ical.vevent.organizer.EMAIL_param
        else:
            sender = ical.vevent.organizer.value.split(':')[1] #workaround for MS
    else:
        sender = "NO SENDER"
    if 'description' in ical.vevent.contents:
        description = ical.vevent.contents['description'][0].value
    else:
        description = "NO DESCRIPTION"
    if 'attendee' in ical.vevent.contents:
        attendees = ical.vevent.contents['attendee']
    else:
        attendees = ""

    print("From:\t\t\t" + sender)
    print("Title:\t\t\t" + summary)
    if hasattr(ical.vevent, 'location'):
        print("Location:\t\t%s" % (ical.vevent.location.value,))
    if hasattr(ical.vevent, 'dtstart'):
        print("Start:\t\t\t%s" % (ical.vevent.dtstart.value,))
    if hasattr(ical.vevent, 'dtend'):
        print("End:\t\t\t%s" % (ical.vevent.dtend.value,))
    print("Attendees:")
    for attendee in attendees:
        if hasattr(attendee,'EMAIL_param') and hasattr(attendee, 'CN_param'):
            print("\t\t\t%s <%s>" % (attendee.CN_param, attendee.EMAIL_param))
        elif hasattr(attendee, 'CN_param'):
            print("\t\t\t%s <%s>" % (attendee.CN_param, attendee.value.split(':')[1]))
        else:
            # workaround for 'mailto:' in email
            print("\t\t\t%s <%s>" % (attendee.value.split(':')[1], attendee.value.split(':')[1]))
    print("Description:")
    print(description)

if __name__=="__main__":

    args, command_base = parser.parse_known_args()

    email_address = None
    accept_decline = 'ACCEPTED'

    invitation = openics(args.icsfile)
    display(invitation)

    if args.v:
        input("Press enter to exit")
        sys.exit(0)

    email_address = args.e
    if args.i:
        accept_decline = get_accept_decline()
    if args.a:
        accept_decline = 'ACCEPTED'
    if args.d:
        accept_decline = 'DECLINED'
    if args.t:
        accept_decline = 'TENTATIVE'

    ans = get_answer(invitation)

    if 'attendee' in invitation.vevent.contents:
        attendees = invitation.vevent.contents['attendee']
    else:
        attendees = ""
    set_accept_state(attendees,accept_decline)
    ans.vevent.add('attendee')
    ans.vevent.attendee_list.pop()
    flag = 1
    for attendee in attendees:
        if hasattr(attendee, 'EMAIL_param'):
            if attendee.EMAIL_param == email_address:
                ans.vevent.attendee_list.append(attendee)
                flag = 0
        else:
            if attendee.value.split(':')[1] == email_address:
                ans.vevent.attendee_list.append(attendee)
                flag = 0
    if flag:
        sys.stderr.write("Seems like you have not been invited to this event!\n")
        sys.exit(1)

    icsfile, tempdir = write_to_tempfile(ans)

    mutt_command = get_mutt_command(command_base, ans, email_address, accept_decline, icsfile)
    mailtext = "%s has %s" % (email_address, accept_decline.lower())
    execute(mutt_command, mailtext.encode('utf-8'))

    os.remove(icsfile)
    os.rmdir(tempdir)
