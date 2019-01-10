from __future__ import absolute_import, print_function
from datetime import datetime
import os
import docusign_esign as docusign
from docusign_esign import AuthenticationApi, TemplatesApi, EnvelopesApi
from docusign_esign.rest import ApiException

# import sys
import csv

# import pandas as pd

wd = 'C:/Users/kevin.alber/Box Sync/Auto FeedBacker/'
os.chdir(wd)

repReports = {}

###########################
# for case list updates and email setting
test_mode = 1
###########################


###########################
# not used currently#
def get_roles():
    fname = 'T3ListNames.csv'
    # fpath = 'C:/Users/kevin.alber/Box Sync/Auto FeedBacker/'
    t3_list = []
    with open(fname, mode='r') as r:
        read = csv.reader(r)
        lists = list(read)
        for line in lists:
            t3_list.append(line[0].lower())

    return t3_list
###########################

# need a CSV with name/email of all t2/t3's
def get_emails():
    repEmails = {}
    # fpath = 'C:/Users/kevin.alber/Box Sync/Auto FeedBacker/'
    fname = 'repEmails.csv'
    with open(fname, mode='r') as l:
        read = csv.reader(l)
        next(read)
        for row in read:
            repEmails[row[0]] = {'email': row[1], 'mgr_email': row[2], 'role': row[3]}
    return repEmails


def updateSentList(caseNumSent):
    fname = 'SentList.csv'
    # fpath = 'C:/Users/kevin.alber/Box Sync/Auto FeedBacker/'
    with open(fname, mode='a', newline='', encoding='utf8') as w:
        wr = csv.writer(w, delimiter=',')
        for i in caseNumSent:
            # print(i)
            wr.writerow([i])


def get_sent_list():
    sent_list = []
    fname = 'SentList.csv'
    # fpath = 'C:/Users/kevin.alber/Box Sync/Auto FeedBacker/'
    with open(fname, mode='r') as w:
        rd = csv.reader(w)
        for case in rd:
            if case[0].startswith('0'):
                sent_list.append(case[0])
            else:
                sent_list.append('0' + case[0])
        print(sent_list)
        return sent_list


def extractReport():
    fname = 'weeklyreport.csv'
    # fpath = 'C:/Users/kevin.alber/Box Sync/Auto FeedBacker/'
    repEmails = get_emails()
    sentList = get_sent_list()
    # print(sentList)

    f = open(fname, encoding="utf-8")
    csv_data = csv.reader(f)
    next(csv_data)
    # roles = get_roles()
    count = 0
    # try:
    for row in csv_data:
        if len(row) > 1:
            caseNum = row[1]
            caseId = row[0]
            editBy = row[4]
            editDate = row[6]
            oldVal = row[8]
            newVal = row[9]
            owner = row[10]
            subject = row[12]
            resolution = row[13]
            topic = row[14]
            subTopic = row[15]
            closedDate = datetime.strptime(row[16], "%m/%d/%Y")
            closedDateStr = closedDate.strftime("%m/%d/%Y")

            # and closedDate > (datetime.now() - timedelta(days=8))

            # if case owner/closer is on our list && (case was escalated to someone in our list OR one of the escalation
            # aliases) && person making the owner change is on our list
            # && case hasn't already been reported on according to our list
            if owner.lower() in repEmails.keys() and (newVal.lower() in repEmails.keys() or newVal in (
            "Support Escalations", "Enterprise Escalations",
            "DevSupport")) and oldVal.lower() in repEmails.keys() and caseNum not in sentList:
                # if owner/closer is tier 3/dev support && they didn't also escalate it && (new owner was set to
                # escalation aliases OR a tier 3/dev) && previous owner isn't one of the support escalation aliases
                # && resolution isn't dupe
                if repEmails[owner.lower()]['role'] in ('t3', 'dev') and oldVal != owner and (
                        newVal in ("Support Escalations", "Enterprise Escalations", "DevSupport") or
                        repEmails[newVal.lower()]['role'] in ('t3', 'dev')) and oldVal not in (
                "Support Escalations", "Enterprise Escalations", "BugBacklog",
                "Enterprise Email Support") and not resolution.startswith("Dupe/Continuation"):
                    # try:
                    # email = repEmails[oldVal]['email']
                    # pass
                    # except:
                    # print ("missing email for " + oldVal)
                    if resolution == "" or not resolution:
                        resolution = "See last case comment"
                    # print(row)

                    # need to test this section, updating dict
                    '''
                    if any(d['CaseNumber'] == caseNum for d in repReports[oldVal]['Cases']): #caseNum in repReports
                    [oldVal]['Cases']:
                        print(repReports)
                        index = next((index for (index, d) in enumerate(repReports[oldVal]['Cases']) if d['CaseNumber']
                         == caseNum), None)
                        if repReports[oldVal]['Cases'][index]["T3RepName"] == "Support Escalations"
                        and repEmails[newVal.lower()]['role'] == 't3':
                            repReports[oldVal]['Cases'][index]["T3RepName"] = newVal
                    else:
                        #fix excel opening issues, it interprets these as functions
                        if resolution.strip().startswith('='):
                            resolution = resolution[1:]
                    '''

                    if oldVal in repReports.keys() and not any(
                            d['CaseNumber'] == caseNum for d in repReports[oldVal]['Cases']):
                        repReports[oldVal]['Cases'].append({"CaseNumber": caseNum,
                            "link": 'https://docusign.my.salesforce.com/console#https%3A%2F%2F'
                            'docusign.my.salesforce.com%2F' + caseId,
                            "DateEscalated": "Escalated: " + editDate + "\nClosed: " + closedDateStr,
                            "T3RepName": owner,
                            "Subject": '(' + topic + ' - ' + subTopic + '): ' + subject,
                            "CaseResolution": resolution})
                        count += 1
                    elif oldVal not in repReports.keys():
                        repReports[oldVal] = {"Email": repEmails[oldVal.lower()]['email'],
                                              "MgrEmail": repEmails[oldVal.lower()]['mgr_email'], "Cases": [
                                {"CaseNumber": caseNum,
                                 "link": 'https://docusign.my.salesforce.com/console#https%3A%2F%2Fdocusign.my.salesforce.com%2F' + caseId,
                                 "DateEscalated": "Escalated: " + editDate + "\nClosed: " + closedDateStr,
                                 "T3RepName": owner, "Subject": '(' + topic + ' - ' + subTopic + '): ' + subject,
                                 "CaseResolution": resolution}]}
                        count += 1
            else:
                # print ("Couldn't find " + owner + ' or ' + newVal + ' or ' + oldVal)
                pass

    # except Exception as e:
    #   print("Broke")

    # print("Should send "+str(count)+" envelopes")
    print(repReports)
    return repReports


# extractReport()

# figure out how many letters I can fit in the box
# this isn't ready yet, it might not work well due to variances in character width
def charLimit(width, lines, entry):
    words = entry.split(' ')
    total = 0
    place = 0
    line = 1
    for wrd in words:
        if line <= lines:
            if (place + len(wrd) + 1) < width:
                total += len(wrd) + 1
                place += len(wrd) + 1
            else:  # go to next line
                line += 1
                place = len(wrd) + 1
    return total


def dsAuth():
    integrator_key = "DOCU-4fcb6bff-70fc-4a48-9fec-2080956f7bcb"
    base_url = "https://na3.docusign.net/restapi"
    oauth_base_url = "account.docusign.com"  # use account.docusign.com for Live/Production
    private_key_filename = "C:/Users/kevin.alber/Box Sync/Auto FeedBacker/docusign_private_key.pem"
    user_id = "d2dc5b01-5dec-45d7-9c3d-b563495dcf3c"  # demo user:"3f6c5781-562e-4604-9ccb-94d782605c26"

    api_client = docusign.ApiClient(base_url)

    # IMPORTANT NOTE:
    # the first time you ask for a JWT access token, you should grant access by making the following call
    # get DocuSign OAuth authorization url:
    # oauth_login_url = api_client.get_jwt_uri(integrator_key, redirect_uri, oauth_base_url)
    # open DocuSign OAuth authorization url in the browser, login and grant access
    # webbrowser.open_new_tab(oauth_login_url)
    # print(oauth_login_url)

    # END OF NOTE

    # configure the ApiClient to asynchronously get an access token and store it
    api_client.configure_jwt_authorization_flow(private_key_filename, oauth_base_url, integrator_key, user_id, 3600)
    return api_client
    # api_client.configure_authorization_flow(integrator_key, client_secret, redirect_uri, oauth_base_url='https://account-d.docusign.com', scope='signature')


# demo templateId: d8e1ec57-a9ee-4518-aa2d-1355126d21d2
def createEnv(repReports=extractReport(), template_id="a07ff1a4-d135-459a-ac45-e5b1f3f5e021", api_client=dsAuth(),
              template_role_name='Rep', template_role_cc='Rep Manager'):
    docusign.configuration.api_client = api_client
    data_labels = {"CaseNum_Label": ["CNum_1", "CNum_2", "CNum_3", "CNum_4", "CNum_5", "CNum_6"],
                   "DateEsc_Label": ["Date_1", "Date_2", "Date_3", "Date_4", "Date_5", "Date_6"],
                   "Subject_Label": ["Subject_1", "Subject_2", "Subject_3", "Subject_4", "Subject_5", "Subject_6"],
                   "CaseRes_Label": ["Resolution_1", "Resolution_2", "Resolution_3", "Resolution_4", "Resolution_5",
                                     "Resolution_6"], "T3": ["T3_1", "T3_2", "T3_3", "T3_4", "T3_5", "T3_6"],
                   "Link": ["Link_1", "Link_2", "Link_3", "Link_4", "Link_5", "Link_6"]}
    # repEmails = get_emails()
    sentListOut = []
    env_count = 0
    # loop through each rep
    for key, val in repReports.items():
        if key != "Bug Backlog" and val[
            'MgrEmail'] != 'n/a':  # 'roarke.mitchell@docusign.com':  # and val['Email'].lower() == 'morgan.jett@docusign.com':
            # create an envelope to be signed
            envelope_definition = docusign.EnvelopeDefinition()
            if val['MgrEmail'] == 'roarke.mitchell@docusign.com' or val['Email'].lower() \
                    in ("marina.mounier@docusign.com","david.linke@docusign.com","jonathan.sammons@docusign.com"):
                envelope_definition.email_subject = 'Tier 3 case escalation feedback - ' \
                                                    'here\'s what happened with your recent previously owned cases'
                envelope_definition.email_blurb = 'Hello ' + key.split(' ')[
                    0] + ', \n\nI\'m sending these out to you tier 3 reps as well now so if another tier 3 or dev ' \
                         'support takes on your cases we can close the loop on these as well. \n\nFeel free to message ' \
                         'me for any suggestions on how to make this program even better. \n\nThanks for reading,' \
                         '\n\nKevin\'s Autofeedbacker bot'
            else:
                envelope_definition.email_subject = 'Tier 3 case escalation feedback - here\'s what happened with ' \
                                                    'your recent escalated cases'
                envelope_definition.email_blurb = 'Hello ' + key.split(' ')[
                    0] + ', \n\nWe\'re starting a new program intended to provide you with a followup on how the cases ' \
                         'you escalated to tier 3 were resolved. The purpose of this is to learn from these cases, ' \
                         'we know you\'re too busy to check up on all of those so we\'re making it easy with these ' \
                         'weekly reports. If you\'re missing any cases we may still be working on them so stay tuned. ' \
                         'Feel free to reach out to the tier 3 rep who closed these cases for more details on what ' \
                         'they did. \n\nFeel free to message Kevin Alber for any suggestions on how to make this ' \
                         'program even better. \n\nThanks for reading,\n\nKevin\'s Autofeedbacker bot'

            # assign template information including ID and role(s)
            envelope_definition.template_id = template_id

            # create manager email notification, need to test

            mgr_emailNotification = docusign.RecipientEmailNotification()
            mgr_emailNotification.email_subject = 'Tier 3 case escalation feedback - here\'s what happened with ' + \
                                                  key.split(' ')[0] + '\'s recent escalated cases'
            mgr_emailNotification.email_body = 'Hello ' + val['MgrEmail'].split('.')[
                0].title() + ', \n\n We\'re starting a new program intended to provide your reps with feedback on cases ' \
                             'that they escalated to tier 3 so they can learn from them and you can better coach them. ' \
                             'Please let me, Kevin Alber, know if you have any feedback about this process. \n\nThanks ' \
                             'for reading,\n\nKevin\'s Autofeedbacker bot'

            # create a template role with a valid template_id and role_name and assign signer info
            t_role = docusign.TemplateRole()
            t_role.role_name = template_role_name
            t_role.name = key
            if test_mode == 1:
                t_role.email = 'kdoc022+reptest@gmail.com'  # for real send
            else:
                t_role.email = val['Email']
                print("Envelope going to rep: " + val['Email'])

            # mgr role on template
            t_role_cc = docusign.TemplateRole()
            t_role_cc.role_name = template_role_cc
            t_role_cc.name = val['MgrEmail'].split('@')[0].replace('.',
                                                                ' ').title()  # might want to just add this to spreadsheet
            if test_mode == 1 or val['MgrEmail'] == 'roarke.mitchell@docusign.com':
                t_role_cc.email = 'kdoc022+mgr@gmail.com'  #
            else:
                t_role_cc.email = val['MgrEmail']
            t_role_cc.email_notification = mgr_emailNotification

            # create a list of template roles and add our newly created role
            # assign template role(s) to the envelope
            envelope_definition.template_roles = [t_role, t_role_cc]

            # send the envelope by setting |status| to "sent". To save as a draft set to "created"
            envelope_definition.status = 'created'

            # notif = Notification()
            # notif.use_account_defaults = "True"

            # envelope_definition.notification = notif
            auth_api = AuthenticationApi()
            envelopes_api = EnvelopesApi()

            # create draft
            try:
                login_info = auth_api.login(api_password='true', include_account_id_guid='true')  ###
                assert login_info is not None
                assert len(login_info.login_accounts) > 0
                login_accounts = login_info.login_accounts
                assert login_accounts[0].account_id is not None

                base_url, _ = login_accounts[0].base_url.split('/v2')
                api_client.host = base_url
                docusign.configuration.api_client = api_client

                envelope_summary = envelopes_api.create_envelope(login_accounts[0].account_id,
                                                                 envelope_definition=envelope_definition)
                assert envelope_summary is not None
                assert envelope_summary.envelope_id is not None
                assert envelope_summary.status == 'created'
                env_count += 1
                # print("Envelope Created Summary: ", end="")
                # pprint(envelope_summary)
            except ApiException as e:
                print("\nException when calling DocuSign API: %s" % e)
                assert e is None  # make the test case fail in case of an API exception

            # get array of tab objects on the current envelope
            tab_array = envelopes_api.get_page_tabs(login_accounts[0].account_id, 1, envelope_summary.envelope_id,
                                                    1).text_tabs

            # iterate through each case as row in envelope
            row = 0
            tab_list = []
            for v in val['Cases']:
                if row < 6:
                    label = 'CaseNum_Label'
                    target_tab = next((x for x in tab_array if x.tab_label == data_labels[label][row]), None)
                    assert target_tab is not None
                    t_col1 = docusign.Text()
                    t_col1.tab_label = data_labels[label][row - 1]
                    t_col1.value = str(v['CaseNumber'])
                    t_col1.tab_id = target_tab.tab_id
                    t_col1.documentId = "1"
                    t_col1.recipientId = "1"
                    t_col1.pageNumber = 1
                    tab_list.append(t_col1)
                    # print(str(v['CaseNumber']))
                    sentListOut.append(v['CaseNumber'])

                    label = 'DateEsc_Label'
                    target_tab = next((x for x in tab_array if x.tab_label == data_labels[label][row]), None)
                    assert target_tab is not None
                    t_col2 = docusign.Text()
                    t_col2.tab_label = data_labels[label][row]
                    t_col2.value = v['DateEscalated']
                    t_col2.tab_id = target_tab.tab_id
                    t_col2.documentId = "1"
                    t_col2.recipientId = "1"
                    t_col2.pageNumber = 1
                    tab_list.append(t_col2)

                    # subMaxChar = charLimit(19, 10, v['Subject'])
                    if len(v['Subject']) > 160:
                        extra = '...'
                    else:
                        extra = ''

                    label = 'Subject_Label'
                    target_tab = next((x for x in tab_array if x.tab_label == data_labels[label][row]), None)
                    assert target_tab is not None
                    t_col3 = docusign.Text()
                    t_col3.tab_label = data_labels[label][row]
                    t_col3.disable_auto_size = 'True'
                    t_col3.height = '200'
                    t_col3.value = v['Subject'][:160] + extra
                    t_col3.tab_id = target_tab.tab_id
                    t_col3.documentId = "1"
                    t_col3.recipientId = "1"
                    t_col3.pageNumber = 1
                    tab_list.append(t_col3)

                    if len(v['CaseResolution']) > 184:
                        extra = '...'
                    else:
                        extra = ''

                    label = 'CaseRes_Label'
                    target_tab = next((x for x in tab_array if x.tab_label == data_labels[label][row]), None)
                    assert target_tab is not None
                    t_col4 = docusign.Text()
                    t_col4.tab_label = data_labels[label][row]
                    t_col4.disable_auto_size = 'True'
                    t_col4.height = '200'
                    t_col4.value = v['CaseResolution'][:184] + extra
                    t_col4.tab_id = target_tab.tab_id
                    t_col4.documentId = "1"
                    t_col4.recipientId = "1"
                    t_col4.pageNumber = 1
                    tab_list.append(t_col4)

                    label = 'T3'
                    target_tab = next((x for x in tab_array if x.tab_label == data_labels[label][row]), None)
                    assert target_tab is not None
                    t_col5 = docusign.Text()
                    t_col5.tab_label = data_labels[label][row]
                    t_col5.value = v['T3RepName']
                    t_col5.tab_id = target_tab.tab_id
                    t_col5.documentId = "1"
                    t_col5.recipientId = "1"
                    t_col5.pageNumber = 1
                    tab_list.append(t_col5)

                    label = 'Link'
                    target_tab = next((x for x in tab_array if x.tab_label == data_labels[label][row]), None)
                    assert target_tab is not None
                    t_col6 = docusign.Text()
                    t_col6.tab_label = "#HREF_" + data_labels[label][row]
                    t_col6.name = v['link']
                    t_col6.value = 'Go to Case'
                    t_col6.tab_id = target_tab.tab_id
                    t_col6.documentId = "1"
                    t_col6.recipientId = "1"
                    t_col6.pageNumber = 1
                    tab_list.append(t_col6)

                    row += 1
                else:
                    print("We're over 6 cases for rep" + val['Email'] + ", dropped case " + str(v['CaseNumber']))

            # now update those tabs using my list of tab definitions, then update status to sent
            tabs = docusign.Tabs()
            tabs.text_tabs = tab_list
            envelopes_api.update_tabs(login_accounts[0].account_id, envelope_summary.envelope_id, 1, tabs=tabs)
            envelope_definition.status = 'sent'
            envelope_summary2 = envelopes_api.update(login_accounts[0].account_id, envelope_summary.envelope_id,
                                                     envelope=envelope_definition)
            if test_mode != 1:
                updateSentList(sentListOut)

            # print("Sent: " +str(key)+ " envId: " + ) # + "EnvelopeSummary: ", end="")
            # pprint(envelope_summary2)
    print("Count: " + str(env_count))


createEnv()
#print(extractReport())
