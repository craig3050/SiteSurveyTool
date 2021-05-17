import requests
import json
import secrets

airtable_api_key = secrets.airtable_api_key
base_key = secrets.base_key
table_name = secrets.table_name


def airtable_download():
    airtable_link = "https://api.airtable.com/v0"
    your_api_key = f"Bearer {airtable_api_key}"

    try:
        with requests.session() as response:
            headers = {
                'Authorization': your_api_key,
            }

            params = (
                ('maxRecords', '1000'),
                ('view', 'Grid view'),
            )
            response = requests.get(f'{airtable_link}/{base_key}/{table_name}',
                                    headers=headers, params=params)

            print(str(response))
            print(str(response.content))
            json_to_print = json.loads(response.content)
            # print(json.dumps(json_to_print, indent=4))

            return json_to_print

    except Exception as e:
        print(e)

def format_airtable_results(airtable_response):
    # Create new dictionary containing all relevant information from Airtable

    for records in airtable_response['records']:
        record_id = records['id']
        try:
            area = records["fields"]["Area"]
            area = area[0]
        except:
            area = "Not recorded"
        try:
            observation_number = records["fields"]["Observation Number"]
        except:
            observation_number = "Not recorded"
        try:
            observation_type = records["fields"]["Observation Type"]
            observation_type = observation_type[0]
        except:
            observation_type = "Not recorded"
        try:
            description = records["fields"]["Description of observation"]
        except:
            description = "Not recorded"
        try:
            created_by = records["fields"]["Created By"]
            print(type(created_by))
        except:
            created_by = "Not recorded"
        try:
            date = records["fields"]["Created time"]
        except:
            date = "Note recorded"
        try:
            image_link = records["fields"]["Attachments"][0]["url"]
        except:
            image_link = "Not recorded"
        try:
            status = records["fields"]["Status"]
        except:
            status = "Open"

        # print(f'{record_id}, {area}, {observation_number}, {observation_type}, {description}, {created_by}, {date}, {status}')
        # print(image_link)
        return record_id, area, observation_number, observation_type, description, created_by, date, status, image_link

def export_to_excel(record):
    #download the picture from the folder and save to correct folder - rename file to the same as the observation number
    #create new row, add the correct details, link to the picture in the folder / reference it
    pass


def export_to_word(record):
    #download picture, insert into correct part of word document
    #see python-docx module - appears to cover everything needed

    pass


if __name__ == '__main__':
    # Download information from Airtable
    airtable_response = airtable_download()
    print(json.dumps(airtable_response, indent=4))

    # returns in format - record_id, area, observation_number, observation_type, description, created_by, date, status, image_link
    formatted_results = format_airtable_results(airtable_response)
    for item in formatted_results:
        print(item)
