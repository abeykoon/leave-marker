import ballerinax/googleapis.gmail;
import ballerina/log;
import ballerina/regex;
import ballerinax/googleapis.sheets as sheets;
import ballerina/lang.array;
import ballerina/time;

const string HALF_DAY_LEAVE_SEARCH_STRING = "half-day leave";
const string LEAVE = "Leave";
const string HALF_DAY_LEAVE = "Half Day";
const string VACAION_EMAIL_START = "Please note that";
const NUMBER_OF_COULMNS_TO_SKIP = 3; //skip first 3 columns for dates
const time:Seconds SECONDS_FOR_DAY = 3600 * 24;

//Gsheet related
const string MEMBER_EMAILS_COLUMN = "C";

configurable OAuth2RefreshTokenGrantConfig gmailOAuthConfig = ?;
configurable OAuth2RefreshTokenGrantConfig gsheetOAuthConfig = ?;
configurable string spreadSheetId = ?;
configurable string workSheetName = ?;

string[] months = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
];

type OAuth2RefreshTokenGrantConfig record {
    string clientId;
    string clientSecret;
    string refreshToken;
    string refreshUrl = "https://oauth2.googleapis.com/token";
};

type LeaveDate record {|
    string date;
    boolean isHalfDay;
|};

type TeamMemberInGSheet record {|
    readonly string email;
    int rowId;
|};

type DateInGsheet record {|
    readonly string date;
    string columnLetter;
    int index;
|};

table<TeamMemberInGSheet> key(email) members = table [];
table<DateInGsheet> key(date) dates = table [];

gmail:Client gmailClient = check new ({
    auth: {
        clientId: gmailOAuthConfig.clientId,
        clientSecret: gmailOAuthConfig.clientSecret,
        refreshToken: gmailOAuthConfig.refreshToken,
        refreshUrl: gmailOAuthConfig.refreshUrl
    }
});

sheets:Client gSheetClient = check new ({
    auth: {
        clientId: gsheetOAuthConfig.clientId,
        clientSecret: gsheetOAuthConfig.clientSecret,
        refreshToken: gsheetOAuthConfig.refreshToken,
        refreshUrl: gsheetOAuthConfig.refreshUrl

    }
});

public function main() returns error? {
    check addColumnsForNextMonthsIfRequired();      //add dates upto next month from current dates to the sheet
    check populateMemberEmailToRowIdTable();
    dates = check populateDateToColumnLetterTable();
    foreach TeamMemberInGSheet teamMember in members {
        do {
            log:printInfo("Checking for email = " + teamMember.email);
            LeaveDate[] latestLeaveDates = check getLatestLeaveDates(teamMember.email);
            if latestLeaveDates.length() == 1 {
                check markSingleLeaveDateInGSheet(teamMember.email, latestLeaveDates[0]);
            } else if latestLeaveDates.length() == 2 {
                check markMultipleLeaveDatesInGSheet(teamMember.email, latestLeaveDates[0], latestLeaveDates[1]);
            } else {
                if (latestLeaveDates.length() != 0) {
                    return error("Unsupported number of elements in latestLeaveDates array");
                }
            }
        } on fail error e {
            log:printError("Error while processing team member " + teamMember.email, e);
        }
    }
}

function addColumnsForNextMonthsIfRequired() returns error? {
    table<DateInGsheet> key(date) existingDates = check populateDateToColumnLetterTable();
    string [] nextMonthDates = getDatesOfFollowingMonth();
    int columnIndexToStartAddingNextMonthDates = existingDates.length() + NUMBER_OF_COULMNS_TO_SKIP;
    boolean areNewColumnsAdded = false;
    foreach string date in nextMonthDates {
        if !existingDates.hasKey(date) {
            if !areNewColumnsAdded {    // add empty columns for new dates. As we add dates accending, we will have to add all dates after this. 
                check addNewColumnsForNewDates(columnIndexToStartAddingNextMonthDates - 1, nextMonthDates, date);
                areNewColumnsAdded = true;
            }
            string columnLetter = indexToColumnName(columnIndexToStartAddingNextMonthDates);
            string A1NotationOfNewDate = columnLetter + "1";    // headers are at first row
            check gSheetClient->setCell(spreadSheetId, workSheetName, A1NotationOfNewDate, date);
            columnIndexToStartAddingNextMonthDates = columnIndexToStartAddingNextMonthDates + 1;
        }
    }
}

function addNewColumnsForNewDates(int columnIndexToStartAddingNextMonthDates, string[] nextMonthDates, string firstAddedDate) returns error? {
    int? indexOfCurrentDate = nextMonthDates.indexOf(firstAddedDate);
    if indexOfCurrentDate is int {
        check gSheetClient-> addColumnsAfterBySheetName(spreadSheetId, workSheetName, columnIndexToStartAddingNextMonthDates, 
        nextMonthDates.length() - indexOfCurrentDate);
    }
}

function getLatestLeaveDates(string teamMemberEmail) returns LeaveDate[]|error {

    string searchQuery = string `from:leave-notification@wso2.com to:${teamMemberEmail} -cc:${teamMemberEmail}`;
    LeaveDate[] leaveDates = [];
    string emailContent = check getLatestEmail(searchQuery);
    if emailContent != "" {
        leaveDates = check extractLeaveDates(emailContent);
    }
    return leaveDates;
}

function getLatestEmail(string searchQuery) returns string|error {
    gmail:MsgSearchFilter searchFilter = {
        q: searchQuery
    };
    string emailBody = "";
    stream<gmail:Message, error?> leaveMessages = check gmailClient->listMessages(searchFilter);
    record {gmail:Message value;}|() leaveMsg = check leaveMessages.next();
    if leaveMsg is record {gmail:Message value;} {
        gmail:MailThread mailThread = check gmailClient->readThread(leaveMsg.value.threadId);
        gmail:Message[]? messages = mailThread.messages;
        if messages is gmail:Message[] {
            foreach gmail:Message message in messages {
                emailBody = message.snippet ?: "";
                if !emailBody.includes(VACAION_EMAIL_START) { //kind reminder, additional comment
                    continue;
                } else {
                    return emailBody;
                }

            }
        } else {
            log:printInfo("This thread has no messages. id = " + mailThread.id);
        }
    } else {
        log:printInfo("No leave emails found for" + searchQuery);
    }
    return emailBody;
}

function extractLeaveDates(string emailBody) returns LeaveDate[]|error {
    //example single leave: Hi all, Please note that Dimitri Abeyeratne will be on 
    //half-day leave (second half) on 08 December 2022. 
    //This email has been sent from an automated system. Please do not reply.

    //example multiple leave:Hi all,
    //Please note that Miraj Abeysekara will be on leave from 21 November 2022 to 02 December 2022.
    LeaveDate[] leaveDates = [];
    string firstSentence = regex:split(emailBody, "\\.")[0];
    boolean isHalfDay = false;
    string startDate = "";
    string endDate = "";

    if firstSentence.includes("from") { //multiple day leave, we only get start and end dates
        int? indexAtFrom = firstSentence.lastIndexOf("from");
        int? indexAtTo = firstSentence.lastIndexOf("to ");  // if we put 'to', strings like october will also be considered. Thus space is required. 
        if (indexAtFrom is int && indexAtTo is int) {
            endDate = firstSentence.substring(indexAtTo + (2 + 1), firstSentence.length());
            string formettedLeaveEndDate = check formatDate(endDate, emailBody);
            boolean isEndDateBeyondToday = check validateLeaveDate(formettedLeaveEndDate);
            if isEndDateBeyondToday {
                startDate = firstSentence.substring(indexAtFrom + (4 + 1), indexAtTo); // 5 = from + space
                LeaveDate leaveStartDate = {
                    date: check formatDate(startDate, emailBody), //check today might be in middle
                    isHalfDay: false
                };
                leaveDates.push(leaveStartDate);
                LeaveDate leaveEndDate = {
                    date: formettedLeaveEndDate,
                    isHalfDay: false
                };
                leaveDates.push(leaveEndDate);
            }
        }
    } else if (firstSentence.includes("on")) { //single day leave
        if firstSentence.includes(HALF_DAY_LEAVE_SEARCH_STRING) {
            isHalfDay = true;
        }
        int? indexOfKeyWord = firstSentence.lastIndexOf("on");
        if indexOfKeyWord is int {
            string leaveDateAsString = firstSentence.substring(indexOfKeyWord + (2 + 1), firstSentence.length()); //ex: 08 December 2022 (3 = on and space)
            string formattedLeaveDate = check formatDate(leaveDateAsString, emailBody); //ex: 08/12/2022
            boolean isDateBeyondToday = check validateLeaveDate(formattedLeaveDate);
            if isDateBeyondToday {
                LeaveDate leaveDate = {
                    date: formattedLeaveDate,
                    isHalfDay: isHalfDay
                };
                leaveDates.push(leaveDate);
            }

        }
    }
    return leaveDates;
}

function markSingleLeaveDateInGSheet(string teamMemberEmail, LeaveDate leaveDate) returns error? {
    string valueToSet;
    if leaveDate.isHalfDay {
        valueToSet = HALF_DAY_LEAVE;
    } else {
        valueToSet = LEAVE;
    }
    string A1NotationOfCell = check getA1Notation(teamMemberEmail, leaveDate.date);
    check gSheetClient->setCell(spreadSheetId, workSheetName, A1NotationOfCell, valueToSet);
}

function markMultipleLeaveDatesInGSheet(string teamMemberEmail, LeaveDate leaveStartDate, LeaveDate leaveEndDate) returns error? {
    string A1NotationOfStartDateCell = check getA1Notation(teamMemberEmail, leaveStartDate.date);
    string A1NotationOfEndDateCell = check getA1Notation(teamMemberEmail, leaveEndDate.date);
    (string|int)[][] leaveDateDataInRange = [];
    (string|int)[] leaveData = [];
    int numberOfDatesOnLeave = check getNumofDatesBeween(leaveStartDate.date, leaveEndDate.date) + 1; //date inclusive, thus add one to the diff
    int counter = 0;
    while counter < numberOfDatesOnLeave {
        leaveData.push(LEAVE);
        counter = counter + 1;
    }
    leaveDateDataInRange.push(leaveData);
    sheets:Range dataRangeToSet = {
        a1Notation: A1NotationOfStartDateCell + ":" + A1NotationOfEndDateCell,
        values: leaveDateDataInRange
    };
    check gSheetClient->setRange(spreadSheetId, workSheetName, dataRangeToSet);
}

function populateMemberEmailToRowIdTable() returns error? {
    sheets:Column columnWithEmails = check gSheetClient->getColumn(spreadSheetId, workSheetName, MEMBER_EMAILS_COLUMN);
    int rowCounter = 1;
    foreach int|string|decimal email in columnWithEmails.values {
        if (rowCounter == 1) {
            rowCounter = rowCounter + 1; //skip header
            continue;
        } else {
            TeamMemberInGSheet member = {
                email: <string>email,
                rowId: rowCounter
            };
            members.add(member);
            rowCounter = rowCounter + 1;
        }
    }
}

function populateDateToColumnLetterTable() returns table<DateInGsheet> key(date) | error {
    sheets:Row rowWithDates = check gSheetClient->getRow(spreadSheetId, workSheetName, 1);
    table<DateInGsheet> key(date) datesInGSheet = table [];
    int columnCounter = 0;
    foreach int|string|decimal date in rowWithDates.values {
        if (columnCounter < NUMBER_OF_COULMNS_TO_SKIP) {
            columnCounter = columnCounter + 1;
            continue;
        } else {
            DateInGsheet dateInGSheet = {
                date: <string>date,
                columnLetter: indexToColumnName(columnCounter),
                index: columnCounter
            };
            datesInGSheet.add(dateInGSheet);
            columnCounter = columnCounter + 1;
        }
    }
    return datesInGSheet;
}

// from https://stackoverflow.com/questions/59401548/how-can-i-convert-an-integer-to-a1-notation
// i is the index from 0
function indexToColumnName(int i) returns string {
    string[] alphabet = [
        "A",
        "B",
        "C",
        "D",
        "E",
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
        "P",
        "Q",
        "R",
        "S",
        "T",
        "U",
        "V",
        "W",
        "X",
        "Y",
        "Z"
    ];
    if (i >= alphabet.length()) {
        decimal x = <decimal>(i / alphabet.length());
        int fx = <int>x.floor();
        return indexToColumnName(fx - 1)
            + indexToColumnName(i % alphabet.length());
    }
    return alphabet[i];
}

function getA1Notation(string teamMemberEmail, string date) returns string|error {
    TeamMemberInGSheet? member = members[teamMemberEmail];
    DateInGsheet? leaveDate = dates[date];
    if (member is TeamMemberInGSheet && leaveDate is DateInGsheet) {
        int rowNum = member.rowId;
        string columnLetter = leaveDate.columnLetter;
        return columnLetter + rowNum.toString();
    } else {
        return error("Cannot generate A1 notation for email = " + teamMemberEmail + " date = " + date);
    }
}

function getNumofDatesBeween(string startDate, string endDate) returns int|error {
    DateInGsheet? startDateInGSheet = dates[startDate];
    DateInGsheet? endDateInGSheet = dates[endDate];
    if (startDateInGSheet is DateInGsheet && endDateInGSheet is DateInGsheet) {
        int datesInBeween = endDateInGSheet.index - startDateInGSheet.index; //index is not from 0, but as we take diff, no issue
        return datesInBeween;
    } else {
        return error("cannot find the number of dates in between startDate = " + startDate + " endDate = " + endDate + ". Are they available in GSheet columns?");
    }
}

//input : 21 November 2022
//output : 21/11/2022
function formatDate(string date, string emailBody) returns string|error {
    string[] parts = regex:split(date, " ");
    int? indexAtMonthArr = array:indexOf(months, parts[1], 0);
    if indexAtMonthArr is int {
        int month = indexAtMonthArr + 1;
        string formattedDate = parts[0] + "/" + month.toString() + "/" + parts[2];
        return formattedDate;
    } else {
        return error("cannot format date " + date + " email = " + emailBody);
    }
}

//Check if date is >= today
function validateLeaveDate(string date) returns boolean|error {
    string[] parts = regex:split(date, "/");
    time:Civil dateAsCivilTime = {
        day: check int:fromString(parts[0]),
        month: check int:fromString(parts[1]),
        year: check int:fromString(parts[2]),
        hour: 23, //get last minute so that we will capture it as a valid date even if we run the program on same day
        minute: 59,
        utcOffset: {hours: 5, minutes: 30}
    };
    time:Seconds utcDiffSeconds = time:utcDiffSeconds(check time:utcFromCivil(dateAsCivilTime), time:utcNow());
    if (utcDiffSeconds > 0d) {
        return true;
    } else {
        return false;
    }
}

function getDatesOfFollowingMonth() returns string [] {
    string[] datesForNextmonth = [];
    time:Utc currentTime = time:utcNow();
    int count = 0;
    while count < 31 {
        time:Utc nextDayUtc = time:utcAddSeconds(currentTime, SECONDS_FOR_DAY);
        time:Civil nextDayCivil = time:utcToCivil(nextDayUtc);
        string dateDateFormatted = nextDayCivil.day < 10 ? string `0${nextDayCivil.day}` : nextDayCivil.day.toString();
        string nextDay = dateDateFormatted + "/" + nextDayCivil.month.toString() + "/" + nextDayCivil.year.toString();    //format = dd/m/yyyy
        datesForNextmonth.push(nextDay);
        currentTime = nextDayUtc;
        count = count + 1;
    }
    return datesForNextmonth; 
}

