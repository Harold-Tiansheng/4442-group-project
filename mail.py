import configparser
import graph
import time
import ms_azure

client = ms_azure.authenticate_client()
#add a message that allows staff to know whether an auto reply is sent or not
sent_message = 'An auto reply has been sent'
unsent_message = "NO auto reply has been sent"

def main():
    print('Initializing\n')

    # Load settings
    config = configparser.ConfigParser()
    config.read(['config.cfg', 'config.dev.cfg'])
    azureSettings = config['azure']

    initialize_graph(azureSettings)

    greet_user()

    current_id = ""

    #graph.create_new_folder()

    while True:
        current_id = list_inbox(current_id)
        #print(current_id)
        time.sleep(60)
        #time between each iteration
        
def initialize_graph(settings: configparser.SectionProxy):
    graph.initialize_graph_for_user_auth(settings)

def greet_user():
    user = graph.get_user()
    print('Hello,', user['displayName'])
    # For Work/school accounts, email is in mail property
    # Personal accounts, email is in userPrincipalName
    print('Email:', user['mail'] or user['userPrincipalName'], '\n')

def list_inbox(current_id):
    message_page = graph.get_inbox()

    # Output each message's details
    i = 0
    for message in message_page['value']:
        if i == 0:
            if message['id'] == current_id:
                break
            else:
                current_id = message['id']
        
        i = i + 1
        try:
            if message['isRead']:
                print("number " + str(i) + " email has been read")
                continue
            reply_address = message['from']['emailAddress']['address']
            documents = [message['body']['content']]
            response = client.extract_key_phrases(documents = documents)[0]
            #print(response.key_phrases)
            if not response.is_error:
                has_send = False
                for phrase in response.key_phrases:
                    if 'Accommodation' in phrase or 'accommodation' in phrase or 'Accomodation' in phrase or 'accomodation' in phrase:
                        graph.do_forward(message['id'], 'h1255740727@gmail.com', sent_message)
                        content = '''
                                <head>Note: this is an auto reply according to email content you provide, hope these information will help you. </head><br>
                            -	<a href="https://www.uwo.ca/univsec/pdf/academic_policies/appeals/Academic%20Accommodation_disabilities.pdf" target="_blank">Academic Accommodation for Students with Disabilities</a><br>
                            -	<a href="https://accessibility.uwo.ca/faculty_staff/policies_programs.html" target="_blank">Accessibility at Western </a><br>
                            -	<a href="https://www.uwo.ca/hr/diversity/accommodate.html" target="_blank">Duty to Accommodate Guidelines </a><br>
                            -	<a href="https://www.uwo.ca/health/" target="_blank">Health and Wellness</a><br>
                            -	<a href="https://www.uwo.ca/coronavirus/" target="_blank">COVID </a><br>
                                <head> If you want help from staff, we will contact with you soon </head><br>
                        '''
                        graph.send_mail('Reply:' + message['subject'], content,reply_address)
                        print("forwarding and auto reply process completed!")
                        has_send = True
                        break
                if not has_send:
                    print("no keyword contain! leave in the mail box")
                    #do nothing
                    #graph.do_forward(message['id'], 'ccornett@uwo.ca')
            else:
                print(response.id, response.error)
        except Exception as err:
            print("Encountered exception. {}".format(err))
    return current_id

main()
