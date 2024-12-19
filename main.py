import json
from extract_quota_usage import extract_quota_usage, run as extract_quota_usage_run
from logging_setup import setup_logging
from file_operations import read_xlsx_to_dict, browse_file, save_file, write_to_template

def calculate_totals(departments, quota_usage_data):
    # Calculate total files and capacity for each department
    results = {}
    for department, sub_departments in departments.items():
        total_files = 0
        total_capacity = 0

        for sub_department in sub_departments:
            # Generate the criteria for the sub-department
            criteria = f"/ifs/{sub_department}".lower()

            # Calculate the total files and total physical capacity
            for quota in quota_usage_data:
                if quota["Path"].lower().startswith(criteria):
                    total_files += quota["Files"]
                    total_capacity += quota["Physical"]

        # Store the results in the dictionary
        results[department] = {
            'total_files': total_files,
            'total_capacity': total_capacity
        }
    return results

def main():
    LOG = setup_logging()
    try:
        # Define departments for LWX
        departments_lwx = {
            'OIT': ['OIT-LW', 'HQAdmins', 'DEPTS', 'OIT'],
            'SECOPS': ['SECOPS'],
            'DPA': ['DPA'],
            'CHS': ['CHS'],
            'DNR': ['DNR'],
            'DOLA': ['DOLA'],
            'DORA': ['DORA'],
            'Public': ['Public'],
            'DOR': ['DOR', 'Revenue'],
            'CST': ['CST'],
            'GOV': ['GOV'],
            'CDA': ['CDA'],
            'HCPF': ['HCPF'],
            'CDOT': ['CDOT', 'CDOTDMZ'],
            'CDEC': ['CDEC', 'CDECHIPAA'],
            'CDPHE': ['CDPHE'],
            'CDLE': ['CDLE'],
            'CDHS': ['CDHS', 'CDHSHIPAA'],
            'Legislative': ['Legislative']
        }

        # Define departments for LWN
        departments_lwn = {
            'OIT': ['OIT'],
            'CDHS': ['CDHS'],
            'DORA': ['DORA'],
            'CDOT': ['CDOT']
        }

        # Prompt the user to select the Ingram Micro cost report file
        ingram_file_path = browse_file("Select the Ingram Micro cost report file", [("Excel files", "*.xlsx")])

        # Load the Ingram Micro cost report
        sheet_name = 'Rating Report'
        data_dict = read_xlsx_to_dict(ingram_file_path, sheet_name, key_column=1, value_start_column=2)
        LWX_Cost = data_dict['lwx400'][20]
        LWN_Cost = data_dict['lwnl400'][20]

        # Prompt the user to select the LWX JSON configuration file
        json_file_path_lwx = browse_file("Select the LWX JSON configuration file", [("JSON files", "*.json")])

        # Load the LWX JSON configuration
        with open(json_file_path_lwx, 'r', encoding='utf8') as json_file:
            json_cfg_lwx = json.load(json_file)

        # Extract LWX quota usage data
        quota_usage_data_lwx = extract_quota_usage(json_cfg_lwx)

        # Prompt the user to select the LWN JSON configuration file
        json_file_path_lwn = browse_file("Select the LWN JSON configuration file", [("JSON files", "*.json")])

        # Load the LWN JSON configuration
        with open(json_file_path_lwn, 'r', encoding='utf8') as json_file:
            json_cfg_lwn = json.load(json_file)

        # Extract LWN quota usage data
        quota_usage_data_lwn = extract_quota_usage(json_cfg_lwn)

        # Calculate totals for LWX and LWN
        data_dict_lwx = calculate_totals(departments_lwx, quota_usage_data_lwx)
        data_dict_lwn = calculate_totals(departments_lwn, quota_usage_data_lwn)

        # Prompt the user to specify the output file name
        output_path = save_file("Save the output file as")

        # Write data to the template
        write_to_template(data_dict_lwx, data_dict_lwn, LWX_Cost, LWN_Cost, output_path, departments_lwx, departments_lwn)
    except ValueError as e:
        LOG.error(e)

if __name__ == "__main__":
    main()