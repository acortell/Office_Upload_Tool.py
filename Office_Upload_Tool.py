import pygsheets
import phonenumbers
import pyap
from urlextract import URLExtract

client = pygsheets.authorize()
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/171M3ZywdrUu9X6TnfN_CyrJVwO6RpNMzFjIYTwTJuEg/edit#gid=349212255")
sheet = spreadsheet.worksheet_by_title("PasteHere")
cells = sheet.get_col(1, include_tailing_empty=False, returnas='cells')  # TODO: Read whole sheet instead of column 1?

final_sheet = [["Name", "Address1", "Address2", "City", "State", "Zip", "Type", "Email", "Phone1", "Phone2", "Fax",
                "Website", "Notes", "Sun", "Mon", "Tue", "Wed", "Thur", "Fri"]]


def search_bold():
    # Remove all rows until the first bold row.
    while 'textFormat' not in list(cells[0].get_json()['userEnteredFormat'].keys()) or \
            'bold' not in list(cells[0].get_json()['userEnteredFormat']['textFormat'].keys()):
        cells.pop(0)

    # Row tracks which row of the final sheet we are working on.
    row = 0

    # For each cell: highlighted ones are a new row for final sheet, others are info for the 2nd cell of same row.
    for c in cells:
        if c.value.strip():
            if 'textFormat' in list(c.get_json()['userEnteredFormat'].keys()) and \
                    'bold' in list(c.get_json()['userEnteredFormat']['textFormat'].keys()):
                if final_sheet[row][12]:
                    final_sheet.append(
                        [c.value, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
                    row += 1
                # If we haven't added any information on the program yet then this line must be more title
                else:
                    final_sheet[row][0] += ' ' + c.value.strip()
                # TODO: Distinguish between consecutive single line offices and titles split over multiple lines
            else:
                if final_sheet[row][12]:
                    final_sheet[row][12] += ' '
                final_sheet[row][12] += c.value.strip()
        else:
            if not final_sheet[row][12]:
                final_sheet[row][12] = '[blank row]'


def search_blanks():
    # Remove all rows until the first non-blank row
    while not cells[0].value.strip():
        cells.pop(0)

    # Row tracks which row of the final sheet we are working on.
    row = 1
    final_sheet.append(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])

    # For each cell: highlighted ones are a new row for final sheet, others are info for the 2nd cell of same row.
    for c in cells:
        if c.value.strip():
            if not final_sheet[row][0]:
                final_sheet[row][0] = c.value
            else:
                if final_sheet[row][12]:
                    final_sheet[row][12] += ' '
                final_sheet[row][12] += c.value.strip()
        else:
            final_sheet.append([c.value, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
            row += 1


def title(text):
    if text:
        ordinals = '1St', '2Nd', '3Rd', '4Th', '5Th', '6Th', '7Th', '8Th', '9Th', '0Th', '1Th', '2Th', '3Th'
        string = text.title()
        for ordinal in ordinals:
            string = string.replace(ordinal, ordinal.lower())
        return string


if '1' in input('If your data has a blank line after each office (and nowhere else), enter 1. Otherwise, make sure your'
                ' office titles are in bold!'):
    search_blanks()
else:
    search_bold()

for v in final_sheet[1:]:
    address = pyap.parse(v[12].upper(), country='US')  # Made Upper because Lower and Title confuse pyap
    if address:
        address_list = [address[0].as_dict()['street_number'], title(address[0].as_dict()['street_name']),
                        title(address[0].as_dict()['street_type']), address[0].as_dict()['route_id'],
                        address[0].as_dict()['post_direction']]
        address1 = [x for x in address_list if x]
        v[1] = ' '.join(address1)
        address2 = [title(address[0].as_dict()['floor']), title(address[0].as_dict()['building_id']),
                    title(address[0].as_dict()['occupancy'])]
        address2 = [x for x in address2 if x]
        v[2] = ' '.join(address2)
        v[3] = title(address[0].as_dict()['city'])
        v[4] = address[0].as_dict()['region1']
        v[5] = address[0].as_dict()['postal_code']

    urls = URLExtract(extract_email=True).find_urls(v[12].lower())

    if urls:
        for url in urls[::-1]:
            if '@' in url:  # This is a simplistic way to find email, a url could also have an @
                v[7] = url  # Overwriting because I have nowhere to store additional urls/emails
            else:
                v[11] = url

    fax = v[12].lower().find('fax')
    if fax > -1:  # Find returns -1 if no instance found
        try:
            match = phonenumbers.PhoneNumberMatcher(v[12][fax:], 'US').next()
            v[12] = v[12][:fax] + v[12][fax:][0:match.start] + '[Redacted]' + v[12][fax:][match.end:]
            v[10] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.RFC3966)[7:]
        except StopIteration:
            pass

    for match in phonenumbers.PhoneNumberMatcher(v[12], "US"):
        if not v[8]:
            v[8] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.RFC3966)[7:]
        elif not v[9]:
            v[9] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.RFC3966)[7:]
        elif not v[10]:
            v[10] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.RFC3966)[7:]
    v[6] = 'service'
    v[14:19] = ['08:00-17:00'] * 5
sheet = spreadsheet.worksheet_by_title("Main")
sheet.update_values("A:S", final_sheet)
