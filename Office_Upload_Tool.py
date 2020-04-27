import pygsheets
import phonenumbers
import pyap
from urlextract import URLExtract
# TODO: Add email finder

client = pygsheets.authorize()
spreadsheet = client.open_by_url("https://docs.google.com/spreadsheets/d/171M3ZywdrUu9X6TnfN_CyrJVwO6RpNMzFjIYTwTJuEg/edit#gid=349212255")
sheet = spreadsheet.worksheet_by_title("PasteHere")
cells = sheet.get_col(1, include_tailing_empty=False, returnas='cells')

# TODO: user input for whether to search by selected cells or use blank lines to delineate. Make 2 functions.

final_sheet = [["Name", "Address1", "Address2", "City", "State", "Zip", "Type", "Email", "Phone1", "Phone2", "Fax",
                "Website", "Notes", "Sun", "Mon", "Tue", "Wed", "Thur", "Fri"]]

while 'textFormat' not in list(cells[0].get_json()['userEnteredFormat'].keys()) or \
      'bold' not in list(cells[0].get_json()['userEnteredFormat']['textFormat'].keys()):
    cells.pop(0)

# Row tracks which row of the final sheet we are working on.
row = 0

# For each cell: highlighted ones are a new row for final sheet, others are information for the 2nd cell of same row.
for c in cells:
    if c.value.strip():
        if 'textFormat' in list(c.get_json()['userEnteredFormat'].keys()) and \
           'bold' in list(c.get_json()['userEnteredFormat']['textFormat'].keys()):
            if final_sheet[row][12]:
                final_sheet.append([c.value, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''])
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

for v in final_sheet[1:]:
    address = pyap.parse(v[12].upper(), country='US')
    if address:
        address_list = [address[0].as_dict()['street_number'], address[0].as_dict()['street_name'].title(),
                        address[0].as_dict()['street_type'].title(), address[0].as_dict()['route_id'],
                        address[0].as_dict()['post_direction']]
        address1 = [x for x in address_list if x]
        v[1] = ' '.join(address1)
        address2 = [address[0].as_dict()['floor'].title(), address[0].as_dict()['building_id'].title(),
                    address[0].as_dict()['occupancy'].title()]
        address2 = [x for x in address2 if x]
        v[2] = ' '.join(address2)
        v[3] = address[0].as_dict()['city'].title()
        v[4] = address[0].as_dict()['region1'].title()
        v[5] = address[0].as_dict()['postal_code']

    urls = URLExtract().find_urls(v[12].lower())
    if urls:
        v[11] = urls[0]  # Find URL from matrix value index 1
    fax = v[12].lower().find('fax')
    if fax > -1:
        try:
            v[10] = [phonenumbers.format_number(x.number, phonenumbers.PhoneNumberFormat.E164) for x in
                     phonenumbers.PhoneNumberMatcher(v[12][fax + 3:], 'US')][0]  # TODO: use consistent format
            v[12] = v[12][:fax] + v[12][fax + 12:]  # TODO: Find a more precise way to do this
        except IndexError:
            pass

    for match in phonenumbers.PhoneNumberMatcher(v[12], "US"):
        if not v[8]:
            v[8] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.NATIONAL)
        elif not v[9]:
            v[9] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.NATIONAL)
        elif not v[10]:
            v[10] = phonenumbers.format_number(match.number, phonenumbers.PhoneNumberFormat.NATIONAL)
    v[6] = 'service'
    v[14:19] = ['08:00-17:00'] * 5
sheet = spreadsheet.worksheet_by_title("Main")
sheet.update_values("A:S", final_sheet)
