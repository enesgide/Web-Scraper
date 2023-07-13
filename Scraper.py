import requests
import pandas as pd
from bs4 import BeautifulSoup
import re
import math

file_name = 'C Services.xlsx'
sheet_n = 'Services'

df = pd.read_excel(file_name, sheet_name=sheet_n)

service_names = df['ServiceName'].tolist()

new_emails = 0

max_entries = 8400
offset = 0

for i in range(max_entries):
    i += offset

    if i % 25 == 0:
        print("Index: " + str(i) + ", Percentage: " + str(math.floor(i/max_entries*100)) + "%")

    try:
        if pd.isna(df.loc[i, "Automated Emails"]):
            name = service_names[i]
            name = name.lower()

            temp_name = ""
            for x in name:
                if x.isalnum() or x == ' ' or x == '-':
                    temp_name += x
            name = temp_name


            name = name.replace(" at ", " ")
            name = name.replace("the ", "")
            name = name.replace(" the ", " ")
            name = name.replace(" by ", " ")
            name = name.replace(" as ", " ")
            name = name.replace(" for ", " ")
            name = name.replace(" on ", " ")
            name = name.replace(" of ", " ")
            name = name.replace(" in ", " ")
            name = name.replace(" and ", " ")
            name = name.replace(" to ", " ")
            name = name.replace(" up ", " ")


            # Request the webpage
            name = name.replace(' ', '-')
            name = name.replace("---", "-")
            name = name.replace("--", "-")

            url = 'https://www.acecqa.gov.au/resources/national-registers/services/' + name
            # print("Url: " + url)
            response = requests.get(url)

            # Parse the HTML content using Beautiful Soup
            soup = BeautifulSoup(response.content, 'html.parser')

            # Find all email addresses on the webpage using regular expression
            email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')

            for link in soup.find_all('a'):
                email = link.get('href')
                if email and email.startswith('mailto:'):
                    email = email[7:]
                    if email_pattern.match(email):
                        # print("Email: " + email)
                        new_emails += 1
                        df.loc[i, "Automated Emails"] = email

            # print("")
    except:
        print("Error at index " + str(i))

writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
df.to_excel(writer, sheet_name=sheet_n, index=False)
writer.close()

print("--- FINISHED ---")
print(str(new_emails) + " new emails added")
