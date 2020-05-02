import pygsheets
import phonenumbers
import pyap
from urlextract import URLExtract

client = pygsheets.authorize()
spreadsheet = client.open_by_url(
    "https://docs.google.com/spreadsheets/d/171M3ZywdrUu9X6TnfN_CyrJVwO6RpNMzFjIYTwTJuEg/edit#gid=349212255")
worksheet = spreadsheet.worksheet_by_title("PasteHere")


def search_bold(in_sheet):
    if len(in_sheet[1]) > 1:
        temp_sheet = [in_sheet.get_row(x + 1, include_tailing_empty=False, returnas='cells') for x in
                      range(len(list(in_sheet))) if ''.join(list(in_sheet)[x])]  # Make list of lists of cells by row
        cells_list = [item for sublist in temp_sheet for item in sublist]  # Turn list of lists into single list
    else:  # If only the first column has values, just get first column
        cells_list = in_sheet.get_col(1, include_tailing_empty=False, returnas='cells')

    out_sheet = [["Name", "Address1", "Address2", "City", "State", "Zip", "Type", "Email", "Phone1", "Phone2", "Fax",
                  "Website", "Notes", "Sun", "Mon", "Tue", "Wed", "Thur", "Fri"]]
    out_row = 0
    for cell in cells_list:
        if cell.value and 'textFormat' in list(cell.get_json()['userEnteredFormat'].keys()) and 'bold' in \
                list(cell.get_json()['userEnteredFormat']['textFormat'].keys()):  # If cell is bold
            if out_sheet[out_row][12]:  # If we've already added info to the current row, make a new one
                out_sheet.append([cell.value, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
                out_row += 1
            else:  # If there's no info on this office yet, assume this cell is more title
                out_sheet[out_row][0] += ' ' + cell.value
        elif cell.value:  # There's something in the cell but it's not bold
            if out_row > 0:  # Ignore values until the first bold value, don't add to header row
                out_sheet[out_row][12] += cell.value + ' '
    return out_sheet


def search_blanks(in_sheet):
    starting_sheet = in_sheet.get_all_values()
    while not starting_sheet[0][0].strip():
        starting_sheet.pop(0)  # Remove all rows until the first row that doesn't begin with a blank cell
    out_sheet_row = 1
    out_sheet = [["Name", "Address1", "Address2", "City", "State", "Zip", "Type", "Email", "Phone1", "Phone2", "Fax",
                  "Website", "Notes", "Sun", "Mon", "Tue", "Wed", "Thur", "Fri"],
                 ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']]
    for row in starting_sheet:
        if ''.join(row).strip():  # If current row is not blank
            if out_sheet[out_sheet_row][0]:  # If current row already has a title, add all to notes
                out_sheet[out_sheet_row][12] += ' '.join([x.strip() for x in row if x.strip()])
            else:  # If current row has no title, first cell is title, add the rest to notes
                out_sheet[out_sheet_row][0] = row[0]
                out_sheet[out_sheet_row][12] += ' '.join([x.strip() for x in row[1:] if x.strip()])
        elif out_sheet[out_sheet_row][0]: # If the current row is blank, start a new row
            out_sheet.append(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
            out_sheet_row += 1
    return out_sheet


def title(text):
    if text:  # Returns nothing if called on none type or empty string or list, avoiding type error
        ordinals = '1St', '2Nd', '3Rd', '4Th', '5Th', '6Th', '7Th', '8Th', '9Th', '0Th', '1Th', '2Th', '3Th'
        string = text.title()
        for ordinal in ordinals:
            string = string.replace(ordinal, ordinal.lower())
        return string


def allot_values():
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


if '1' in input('If your data has a blank line after each office (and nowhere else), enter 1. Otherwise, make sure your'
                ' office titles are in bold!'):
    final_sheet = search_blanks(worksheet)
else:
    final_sheet = search_bold(worksheet)

allot_values()

worksheet = spreadsheet.worksheet_by_title("Main")
worksheet.update_values("A:S", final_sheet)
