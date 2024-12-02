// This script is executed periodically to sync new submitted registrant forms with account google contacts.

// Operation:
// 1. Check for new emails since last execution time for matching emails
// 2. Extract new registrant data from form in email body
// 3. Create new Google Contact for registrant if doesn't exist

function updateLastExecutionDate(): void {
    PropertiesService.getScriptProperties()
        .setProperty(Settings.LAST_EXECUTION_DATE_PROP_NAME, new Date().toString());
}
    
function getLastExecutionDate(): Date {
    const lastDateString = PropertiesService.getScriptProperties()
                            .getProperty(Settings.LAST_EXECUTION_DATE_PROP_NAME);
    return lastDateString && new Date(lastDateString);
}

function getLookbackStart(lastTriggerDate: Date): Date {
    const maxLookback = new Date(new Date().setDate(new Date().getDate() - Settings.MAX_LOOKBACK_DAYS));

    return lastTriggerDate < maxLookback ? maxLookback : lastTriggerDate;
}


function onTimeTrigger() {
    console.log(`Time trigger started`);
    onTrigger();
}

function onTrigger() {
    const lastTriggerDate = getLastExecutionDate();
    console.log(`Last execution timestamp: ${lastTriggerDate.toString()}`);

    const start = getLookbackStart(lastTriggerDate);
    console.log(`Looking back at new registrants after: ${start.toString()}`);


}