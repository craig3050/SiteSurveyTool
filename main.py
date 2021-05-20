import requests
import json
import secrets
import datetime
import os
from exception_decor import exception
from exception_logger import logger
from docx import Document
from PIL import Image, ImageOps

airtable_api_key = secrets.airtable_api_key
base_key = secrets.base_key
table_name = secrets.table_name

@exception(logger)
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

@exception(logger)
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
    try:
        entry_dict['location'] = records["fields"]["Location"]
    except:
        entry_dict['location'] = "Not recorded"
    return entry_dict

def export_to_excel(record):
    #download the picture from the folder and save to correct folder - rename file to the same as the observation number
    #create new row, add the correct details, link to the picture in the folder / reference it
    pass


def export_to_word(formatted_results, image_link):
    if os.path.isfile(f'SiteReport.docx'):
        print("File already exists")
        doc = Document("SiteReport.docx")
    else:
        doc = Document("Template.docx")

    #print(doc.tables)
    # print("Retrieved value: " + doc.tables[0].cell(0, 0).text)
    # print("Retrieved value: " + doc.tables[1].cell(0, 0).text)
    # print("Retrieved value: " + doc.tables[2].cell(0, 0).text)
    # for item in formatted_results:
    #     print(item, formatted_results[item])
    row_count = len(doc.tables[2].rows)
    # print(f'Row Count: {row_count}')
    text_for_table = f"""
Area: {str(formatted_results['area'])}
Location: {str(formatted_results['location'])}
Observation Type: {str(formatted_results['observation_type'])}
Observation Category: {str(formatted_results['observation_category'])}
Description:\n {str(formatted_results['description'])}
                     """
    doc.tables[2].cell(row_count-1, 0).text = str(formatted_results['observation_number'])
    doc.tables[2].cell(row_count-1, 1).text = str(text_for_table)
    paragraph = doc.tables[2].cell(row_count-1, 2).paragraphs[0]
    #paragraph.text = image_link
    # print(image_link)
    try:
        #paragraph.add_picture(image_link, width=530)
        run = paragraph.add_run()
        run.add_picture(image_link, width = 3000000)
    except Exception as e:
        print(e)
    #doc.tables[2].cell(row_count-1, 2).add_picture(image_data, width=Cm(10))
    doc.tables[2].add_row() #ADD ROW HERE
    doc.save("SiteReport.docx")

@exception(logger)
def download_picture(picture_link, observation_number):
    directory_name = 'Pictures'
    try:
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
            try:
                image_pil = Image.open(f'{directory_name}/{observation_number}.jpg').convert('RGB')
                image_pil = ImageOps.exif_transpose(image_pil)
                print(type(image_pil))
                image_pil.save(f'{directory_name}/{observation_number}.jpg')
            except:
                "Cannot transpose image"
    except:
        print('No image to download')

    return f'{directory_name}/{observation_number}.jpg'


if __name__ == '__main__':
    # Download information from Airtable
    airtable_response = airtable_download()
    print(json.dumps(airtable_response, indent=4))

    # Send download data format into correct sections
    for records in airtable_response['records']:
        formatted_results = format_airtable_results(records)
        image_link = download_picture(formatted_results['image_link'], formatted_results['observation_number'])
        export_to_word(formatted_results, image_link)

