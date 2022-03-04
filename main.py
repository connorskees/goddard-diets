import pandas as pd

all_attendees = pd.read_excel("/Users/user/Downloads/All guests_Connor.xlsx")
attendees_with_guests = all_attendees[all_attendees['haveguest'] == 'X']

print(len(attendees_with_guests))

all_guests = attendees_with_guests[[
    'tablenumber',  'first', 'last', 'descrip', 'host', 'guestfirst', 'guestlast', 'guestdietary']]

all_guests['descrip'] = 'Guest of ' + \
    all_guests['first'] + ' ' + all_guests['last']

all_guests = all_guests.drop(columns=['first', 'last'])

all_guests = all_guests.rename(
    columns={'guestfirst': 'first', 'guestlast': 'last', 'guestdietary': 'dietary'})

attendees_with_allergies = all_attendees[[
    'tablenumber',  'descrip', 'host', 'first', 'last', 'dietary']][all_attendees['dietary'].notnull()]

attendees_with_allergies = attendees_with_allergies.append(
    all_guests[all_guests['dietary'].notnull()])

print(len(all_guests[all_guests['dietary'].notnull()]))

# print(
#     (attendees_with_allergies))

# attendees_with_allergies.to_excel(
#     "hilton-goddard-dietary-2022-03-02.xlsx", index=False)
