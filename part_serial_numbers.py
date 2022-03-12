def XLSExport(Rows, SheetName, FileName):
    from openpyxl import Workbook
    wb = Workbook()

    ws = wb.active
    ws.title = SheetName
    # ws = wb.create_sheet(SheetName)
    for x in Rows:
        ws.append(x)

    wb.save(FileName)
    
def card_mda_detail_parser(data_to_parse):
    ttp_template = template_card_mda_detail_parser

    parser = ttp(data=data_to_parse, template=ttp_template)
    parser.parse()

    # print result in JSON format
    results = parser.result(format='json')[0]
    #print(results)

    #converting str to json. 
    result = json.loads(results)
    
    return(result)

ExcelExport = [["Node_IP","Card_MDA_No", "Part_Numbers", "Serial_Numbers"]]
#ExcelExport = [["CI", "IP", "Global Network", "Global Next Hop", "TAC"]]

#list of IPs to be checked. 
ip_list = ["172.29.6.164", "172.29.6.174", "172.29.6.137", "172.29.6.158", "172.29.6.121", "172.29.6.129", "172.29.6.124"]

for ip in ip_list:

    targetnode = {
    'device_type': 'nokia_sros',
    'ip': ip,
    'username': 'admin',
    'password': 'admin',
    'port': 22,
    }

    remote_connect = ConnectHandler(**targetnode)
    remote_connect.send_command("environment no more\n")
    card_output = remote_connect.send_command("show card detail")
    mda_output = remote_connect.send_command("show mda detail") 

    parsed_card_detail_output = card_mda_detail_parser(card_output)
    parsed_mda_detail_output = card_mda_detail_parser(mda_output)
    print("#####################################################################\n")
    print(f"Node IP {ip} is being checked... Please see details as following...")
    print("---------------------------------------------------------------------\n")
    time.sleep(4)

    with open("clishow_parser_outputs\getting_part_serial_number.txt", "a") as f:
        f.write(f"\nNode IP : {ip}\n\n")
    print("Card Detail for this node is being checked...\n")
    for card_elo in parsed_card_detail_output[0]['Card_No']:
        #print(card_elo)
        if 'Card_Detail' in card_elo:
            print(f"See Part and Serial number informations for card {card_elo['Card_ID']} : Part Number : {card_elo['Card_Detail']['Part_Number']} Serial Number : {card_elo['Card_Detail']['Serial_Number']}")
            with open("clishow_parser_outputs\getting_part_serial_number.txt", "a") as f:
                f.write(f"Card ID : {card_elo['Card_ID']} --> Part Number : {card_elo['Card_Detail']['Part_Number']} --> Serial Number : {card_elo['Card_Detail']['Serial_Number']}\n")
            ExcelExport.append([ip,card_elo['Card_ID'],card_elo['Card_Detail']['Part_Number'],card_elo['Card_Detail']['Serial_Number']])
        elif 'Card_Detail' not in card_elo:
            print(f"It looks card id {card_elo['Card_ID']} is operationally down state.")
    print("\nCard Detail for this node has been done.\n")
    time.sleep(2)
    print("MDA Detail for this node is being checked...\n")
    for mda_elo in parsed_mda_detail_output[0]['MDA_No']:
        #print(card_elo)
        if 'MDA_Detail' in mda_elo:
            print(f"See Part and Serial number informations for mda {mda_elo['MDA_ID']} : Part Number : {mda_elo['MDA_Detail']['Part_Number']} Serial Number : {mda_elo['MDA_Detail']['Serial_Number']}")
            with open("clishow_parser_outputs\getting_part_serial_number.txt", "a") as f:
                f.write(f"MDA ID : {mda_elo['MDA_ID']} --> Part Number : {mda_elo['MDA_Detail']['Part_Number']} --> Serial Number : {mda_elo['MDA_Detail']['Serial_Number']}\n")
            ExcelExport.append([ip,mda_elo['MDA_ID'],mda_elo['MDA_Detail']['Part_Number'],mda_elo['MDA_Detail']['Serial_Number']])
            #ExcelExport.append([HostName, x, words14, IP2, TAC])
        elif 'MDA_Detail' not in mda_elo:
            print(f"It looks mda id {mda_elo['MDA_ID']} is operationally down state.")
    with open("clishow_parser_outputs\getting_part_serial_number.txt", "a") as f:
        f.write(f"\n##########################################################################\n")    
    print("\nMDA Detail for this node has been done.\n")

    XLSExport(ExcelExport, "INFORMATION", "LAB_PART_SERIAL_NUMBERS.xlsx") ## xlsx file has been created. 

    
