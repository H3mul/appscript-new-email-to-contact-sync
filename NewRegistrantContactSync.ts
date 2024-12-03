// This script is executed periodically to sync new submitted registrant forms with account google contacts.

// Operation:
// 1. Check for new emails since last execution time for matching emails
// 2. Extract new registrant data from form in email body
// 3. Create new Google Contact for registrant if doesn't exist

const RE_NEWLINE = "\\r\\n";

function updateLastLookbackDate(date: Date): void {
    PropertiesService.getScriptProperties()
        .setProperty(Settings.LAST_LOOKBACK_DATE_PROP_NAME, date.toString());
}
    
function getLastLookbackDate(): Date {
    const lastDateString = PropertiesService.getScriptProperties()
                            .getProperty(Settings.LAST_LOOKBACK_DATE_PROP_NAME);
    return lastDateString && new Date(lastDateString);
}

function getLookbackStart(): Date {
    const lastDate = getLastLookbackDate();
    const maxLookback = new Date(new Date().setDate(new Date().getDate() - Settings.MAX_LOOKBACK_DAYS));

    return lastDate < maxLookback ? maxLookback : lastDate;
}

function getRequiredProp(propName: string) {
    const value = PropertiesService.getScriptProperties().getProperty(propName);
    if (!value) {
        throw new Error(`Required script property not set: ${propName}`);
    }
    return value;
}

function getTargetEmailsAfter(start: Date): GoogleAppsScript.Gmail.GmailMessage[] {
    const labelName = getRequiredProp(Settings.EMAIL_LABEL_PROP_NAME);
    const targetSubject = new RegExp(getRequiredProp(Settings.EMAIL_SUBJECT_PROP_NAME),'i');
    const label = GmailApp.getUserLabelByName(labelName);
    
    if (!label) {
        throw new Error(`Label not found for name: ${labelName}`);
    }

    const targetMessages = [];

    while (targetMessages.length < Settings.MAX_MESSAGE_PROCESSING_COUNT) {

        // Only interested in first message of threads
        const newMessages = label.getThreads(0, Settings.MAX_THREAD_COUNT_REQUEST)
            .filter(t => t.getMessageCount() > 0)
            .map(t => t.getMessages()[0]);

        // Add new message matching subject and after start date
        targetMessages.push(
            ...newMessages.filter(m => m.getDate() > start && targetSubject.test(m.getSubject()))
        );

        if (!newMessages.length || newMessages[newMessages.length - 1].getDate() < start) {
            break;
        }
    }
    return targetMessages;
}

type Contact = {
    givenName: string;
    lastName: string;    
    email: string;
    phoneNumber: string;
    address: string;
    studentName?: string;
}

function extractFormField(form: string[], field: string, lines = 1): string {
    const re = new RegExp(field, 'i');

    let headingIndex = -1;
    form.some((l, i) => {
        if (re.test(l)) {
            headingIndex = i;
            return true;
        }
    });

    if (headingIndex === -1 || form.length < headingIndex + 1 + lines) {
        throw new Error(`Failed to extract form field: ${field}`);
    }

    return form.slice(headingIndex + 1, headingIndex + 1 + lines).join(' ').trim();
}

function extractContactFromMessage(message: GoogleAppsScript.Gmail.GmailMessage): Contact {
    const formBody = message.getPlainBody().split('\n');

    const [ givenName, lastName ] = extractFormField(formBody, "Your Name").split(' ');
    return {
        givenName, lastName,
        email: extractFormField(formBody, "Email"),
        phoneNumber: extractFormField(formBody, "Mobile Phone"),
        address: extractFormField(formBody, "Address", 3),
        studentName: extractFormField(formBody, "Student's First and Last Name"),
    }
}

function createMissingContacts(contacts: Contact[]) {
    const existingContacts = People.People.Connections.list('people/me', { personFields: 'emailAddresses' });

    const newContacts = contacts.filter(c => 
        existingContacts?.connections.some((connection) =>
            connection.emailAddresses && connection.emailAddresses.some((emailAddress) => emailAddress.value === c.email)
    ));

    newContacts.forEach(c => {
        People.People.createContact({
            names: [{ givenName: c.givenName, familyName: c.lastName }],
            emailAddresses: [{ value: c.email }],
            phoneNumbers: [{ value: c.phoneNumber }],
            addresses: [{ streetAddress: c.address }],
            userDefined: [{ key: "Student's Name", value: c.studentName }]
        })            
    });

    return newContacts;
}

function onTimeTrigger() {
    console.log(`Time trigger started`);
    onTrigger();
}

function onTrigger() {
    console.log(`Starting New Registrant Contact Sync script with properties:`);
    console.log(`   ${JSON.stringify(PropertiesService.getScriptProperties().getProperties())}`);

    const start = getLookbackStart();
    console.log(`Looking back at new emails since: ${start.toString()}`);

    const messages = getTargetEmailsAfter(start);
    console.log(`Found ${messages.length} new target messages`);

    if (messages.length) {
        updateLastLookbackDate(new Date(messages[0].getDate().toString()))
    }

    const contacts = messages.map(m => extractContactFromMessage(m));
    const createdContacts = createMissingContacts(contacts);

    console.log(`Added ${createdContacts.length} new contacts ${createdContacts.map(c => c.email).join(', ')}`);
}