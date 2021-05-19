import requests
import json
import secrets
import datetime
import os

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

def format_airtable_results(records):
    # Create new dictionary containing all relevant information from Airtable
    entry_dict = {}

    entry_dict['record_id'] = records['id']
    try:
        area = records["fields"]["Area"]
        entry_dict['area'] = area[0]
    except:
        entry_dict['area'] = "Not recorded"
    try:
        entry_dict['observation_number'] = records["fields"]["Observation Number"]
    except:
        entry_dict['observation_number'] = "Not recorded"
    try:
        observation_type = records["fields"]["Observation Type"]
        entry_dict['observation_type'] = observation_type[0]
    except:
        entry_dict['observation_type'] = "Not recorded"
    try:
        entry_dict['description'] = records["fields"]["Description of observation"]
    except:
        entry_dict['description'] = "Not recorded"
    try:
        entry_dict['created_by'] = records["fields"]["Created By"]
    except:
        entry_dict['created_by'] = "Not recorded"
    try:
        date_object = datetime.datetime.strptime(records["fields"]["Created time"], '%Y-%m-%dT%H:%M:%S.%fZ')
        entry_dict['date'] = date_object.strftime("%d.%m.%Y")
    except Exception as e:
        print(e)
        entry_dict['date'] = "Not recorded"
    try:
        entry_dict['image_link'] = records["fields"]["Attachments"][0]["url"]
    except:
        entry_dict['image_link'] = "Not recorded"
    try:
        entry_dict['status'] = records["fields"]["Status"]
    except:
        entry_dict['status'] = "Open"
    try:
        entry_dict['observation_category'] = records["fields"]["Observation Category"]
    except:
        entry_dict['observation_category'] = "Not recorded"

    return entry_dict

def export_to_excel(record):
    #download the picture from the folder and save to correct folder - rename file to the same as the observation number
    #create new row, add the correct details, link to the picture in the folder / reference it
    pass


def export_to_word(record):
    #download picture, insert into correct part of word document
    #see python-docx module - appears to cover everything needed

    pass

def download_picture(picture_link, observation_number):
    try:
        directory_name = 'Pictures'
        try:
            os.makedirs(directory_name)
        except:
            print("Directory already exists")

        if os.path.isfile(f'{directory_name}/{observation_number}.jpg'):
            print("File already exists")
        else:
            print("File does not exist")
            img_data = requests.get(picture_link).content
            try:
                with open(f'{directory_name}/{observation_number}.jpg', 'wb') as handler:
                    handler.write(img_data)
            except Exception as e:
                print(e)

    except:
        return 0



if __name__ == '__main__':
    # Download information from Airtable
    airtable_response = airtable_download()
    print(json.dumps(airtable_response, indent=4))

    # Send download data format into correct sections
    for records in airtable_response['records']:
        formatted_results = format_airtable_results(records)
        for item in formatted_results:
            print(item, formatted_results[item])
        print(formatted_results['image_link'])
        image_data = download_picture(formatted_results['image_link'], formatted_results['observation_number'])

