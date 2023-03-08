import os
import shutil
import glob
import requests
import pandas as pd



# Get the request data
def my_request(zone_range):
    # https://www.ups.com/media/us/currentrates/zone-csv/200.xls
    # print(zone_range.split("-")[0][:3])
    url = f'https://www.ups.com/media/us/currentrates/zone-csv/{zone_range.split("-")[0][:3]}.xls'
    # print(url)
    response = requests.get(url)
    return response


# Writing data to a file
def write_data_to_file(filename, regime,data):
    with open(filename, regime) as f:
        f.write(data)


# Check that the downloaded file matches the expected range
def check_ranges(data, zone_range):
    downloaded_range_start = pd.read_excel(data, header=None, nrows=5).iat[4, 0].split(" ")[6].replace("-", "")[:4] + "0"
    downloaded_range_end = pd.read_excel(data, header=None, nrows=5).iat[4, 0].split(" ")[8].replace("-", "")[:5]
    print(f'ZRS -- {zone_range.split("-")[0]}, DRS -- {downloaded_range_start}, ZRE -- {zone_range.split("-")[1]}, DRE -- {downloaded_range_end}')

    new_zone_range = []
    if downloaded_range_start == zone_range.split("-")[0] and downloaded_range_end == zone_range.split("-")[1]:
        new_zone_range.append(f'{downloaded_range_start}-{downloaded_range_end}')
        return [True, new_zone_range]
    else:
        new_zone_range.append(f'{downloaded_range_start}-{downloaded_range_end}')
        prefix_length = len(downloaded_range_end) - len(downloaded_range_end.lstrip('0'))
        new_zone_range.append(f'{(str(prefix_length*"0") + str(int(downloaded_range_end)+1))}-{zone_range.split("-")[1]}')
        return [False, new_zone_range]
         
    

# Convert xls files to xlsx files
def convert_xls_to_xlsx(input_dir):
    output_dir = os.path.join(os.path.dirname(input_dir), "UPS_domestic_zones_xlsx")
    # print(output_dir)
    
    # Check if output directory exists, recreate it if necessary
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)
    
    # Find all .xls files in input folder
    xls_files = glob.glob(os.path.join(input_dir, '*.xls'))
    
    for xls_file in xls_files:
        # Read .xls file into pandas dataframe
        xls_df = pd.read_excel(xls_file, engine='openpyxl')
        
        # Get filename without extension
        filename = os.path.splitext(os.path.basename(xls_file))[0]
        
        # Save dataframe to .xlsx file in output folder
        xlsx_file = os.path.join(output_dir, f'{filename}.xlsx')
        xls_df.to_excel(xlsx_file, index=False)



# Download UPS domestic zone charts
def download_UPS_domestic_zone_charts(zone_ranges):   
    corrected_zone_ranges = []

    for zone_range in zone_ranges:
        print(zone_range)
        response = my_request(zone_range)
        info_check = check_ranges(response.content, zone_range=zone_range)
        # print(info_check)
        if info_check[0]:
            corrected_zone_ranges.append(info_check[1][0])
            filename = f'UPS_domestic_zones/{info_check[1][0]}.xls'
            write_data_to_file(filename=filename, regime="wb", data=response.content)
            # print(1)
        
        if info_check[0] == False:
            for range in info_check[1]:
                response = my_request(range)
                info_check = check_ranges(response.content, zone_range=range)
                corrected_zone_ranges.append(range)
                filename = f'UPS_domestic_zones/{range}.xls'
                write_data_to_file(filename=filename, regime="wb", data=response.content)
                # print(2)
    
    return corrected_zone_ranges




# Download UPS domestic zone charts
# def test():
#     corrected_zone_ranges = []
#     for zone_range in zone_ranges:
#         response = my_request(zone_range)
#         filename = f'UPS_domestic_zones/{zone_range}.xls'
        
#         # Check that the downloaded file matches the expected range
#         downloaded_range_start = pd.read_excel(response.content, header=None, nrows=5).iat[4, 0].split(" ")[6].replace("-", "")[:4] + "0"
#         downloaded_range_end = pd.read_excel(response.content, header=None, nrows=5).iat[4, 0].split(" ")[8].replace("-", "")[:5]
#         print(f'ZRS -- {zone_range.split("-")[0]}, DRS -- {downloaded_range_start}, ZRE -- {zone_range.split("-")[1]}, DRE -- {downloaded_range_end}')

#         if downloaded_range_start == zone_range.split("-")[0] and downloaded_range_end == zone_range.split("-")[1]:
#             # print(f'{downloaded_range_start}-{downloaded_range_end}')
#             corrected_zone_ranges.append(f'{downloaded_range_start}-{downloaded_range_end}')
#             write_data_to_file(filename=filename, regime="wb", data=response.content)

#         else:
#             print(f"The downloaded file={f'{downloaded_range_start}-{downloaded_range_end}'} does not match the expected {zone_range}.")
#             print(f'ZRE -- {zone_range.split("-")[1]}, DRE -- {downloaded_range_end}')
#             prefix_length = len(downloaded_range_end) - len(downloaded_range_end.lstrip('0'))
#             new_zone_range = f'{(str(prefix_length*"0") + str(int(downloaded_range_end)+1))}-{zone_range.split("-")[1]}'
#             # print(new_zone_range)
#             corrected_zone_ranges.append(new_zone_range)
#             response = my_request(new_zone_range)
#             filename = f'UPS_domestic_zones/{new_zone_range}.xls'
#             # print(filename)
#             write_data_to_file(filename=filename, regime="wb", data=response.content)



if __name__ == "__main__":

    # zone_ranges = ['00500-00599', '01000-01199']

    # Read zone ranges from Carriers zone ranges (1).xlsx
    zones_df = pd.read_excel('Carriers zone ranges (1).xlsx', sheet_name='UPS zip ranges')
    zone_ranges = [f"{int(row['zip from']):05d}-{int(row['zip to']):05d}" for _, row in zones_df.iterrows()]
    # print(zone_ranges)
    

    # Delete existing UPS_domestic_zones directory and create new one
    shutil.rmtree('UPS_domestic_zones', ignore_errors=True)
    os.makedirs('UPS_domestic_zones')

    corrected_zone_ranges = download_UPS_domestic_zone_charts(zone_ranges)
    # Read the excel file
    # df = pd.read_excel('Carriers zone ranges (1).xlsx', sheet_name='UPS zip ranges')
    # print(df)
    df = pd.DataFrame({'UPS zone ranges': corrected_zone_ranges})
    df.to_excel("output.xlsx", sheet_name='UPS zip ranges', index=False)

    convert_xls_to_xlsx("UPS_domestic_zones")

