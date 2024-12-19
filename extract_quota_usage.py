import sys
import os
import json
import logging
import optparse

# Global logging object
LOG = None
DEFAULT_LOG_FORMAT = '%(asctime)s - %(module)s|%(funcName)s - %(levelname)s [%(lineno)d] %(message)s'

def add_parser_options(parser):
    # Add command line options for JSON file and encoding
    parser.add_option("--json_file", "-j",
                      default=None,
                      help="Path and file name of the JSON config data.\n"
                           "(Default: None)")
    parser.add_option("--json_encoding",
                      default='utf8',
                      help="Set the character encoding when opening the JSON file. This"
                           " may be required if you get an error like:"
                           " UnicodeDecodeError: 'utf-8' codec can't decode byte...")

def extract_quota_usage(json_cfg):
    # Extracting relevant fields for "Quota Usage" from the JSON config
    quota_usage_data = []
    quota_usage = json_cfg.get('stats', {}).get('smartquotas', {}).get('usage', [])
    if not quota_usage:
        print("No quota usage data found.")
    else:
        print(f"Found {len(quota_usage)} quota usage entries.")
    for quota in quota_usage:
        quota_usage_data.append({
            "Path": quota.get("path", ""),
            "Type": quota.get("type", ""),
            "Linked": quota.get("linked", ""),
            "Persona": quota.get("name", ""),
            "Files": quota.get("inodes", 0),
            "Physical": quota.get("physical", 0),
            "FS Physical": quota.get("physical_data", 0),
            "FS Logical": quota.get("logical", 0),
            "App Logical": quota.get("applogical", 0),
            "Shadow Logical": quota.get("shadow_refs", 0),
            "Protection": quota.get("physical_protection", 0),
            "Reduction Ratio": quota.get("reduction_ratio", ""),
            "Efficiency Ratio": quota.get("efficiency_ratio", "")
        })
    return quota_usage_data

def run():
    global LOG
    # Create our command line parser
    parser = optparse.OptionParser()
    add_parser_options(parser)
    (options, args) = parser.parse_args(sys.argv[1:])

    # Setup logging
    LOG = logging.getLogger()
    LOG.setLevel(logging.DEBUG)
    log_handler = logging.StreamHandler()
    log_handler.setFormatter(logging.Formatter(DEFAULT_LOG_FORMAT))
    LOG.addHandler(log_handler)

    if not options.json_file:
        parser.print_help()
        parser.error('You must specify an input JSON configuration file')

    try:
        # Load the JSON configuration file
        with open(options.json_file, encoding=options.json_encoding) as json_data:
            json_cfg = json.load(json_data)
            print("Loaded JSON configuration successfully.")
    except UnicodeDecodeError as e:
        # Handle encoding errors
        LOG.error(
            "JSON input file could not be read using UTF-8 character encoding. Try running with"
            " --json_encoding=latin-1 or --json_encoding=ascii command line options."
            " Error starting at byte: %s" % (e.start)
        )
        sys.exit(7)
    except:
        # Handle other errors
        LOG.exception("Could not load or parse: %s" % options.json_file)
        sys.exit(2)

    # Extract quota usage data
    quota_usage = extract_quota_usage(json_cfg)
    
    # Print the extracted quota usage data to the console
    print(json.dumps(quota_usage, ensure_ascii=False, indent=4))

if __name__ == "__main__":
    run()
