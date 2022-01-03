import json
offset = 0
record_list = []
while offset is not None:
    with open('json_output2.json', 'r') as f:
        item_to_print = json.load(f)
        for item in item_to_print['records']:
            record_list.append(json.dumps(item))
        # print(item_to_print['records'])
        # print(type(item_to_print))
        offset = item_to_print['offset']
        print("====================================================================")
        print(record_list)

print(offset)
print(record_list)


#     # for item in item_to_print:
#     #     print(item)
#     # for records in item_to_print['records']:
#     #     print(records)
#     with open('json_output3.json', 'w') as file:
#         json_file_to_write = item_to_print
#         for item in item_to_print['records']:
#             json_file_to_write.append(item)
#             file.write(json.dumps(json_file_to_write))
#
# print(json_file_to_write)
# print(item_to_print['offset'])

