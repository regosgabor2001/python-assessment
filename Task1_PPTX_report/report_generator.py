import argparse
import json

def read_config_file(config_file):
    with open(config_file, 'r') as f:
        config = json.load(f)
    return config

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Generate a report in pptx format based on a configuration file')
    parser.add_argument('sample.json', help='path to the configuration file')
    args = parser.parse_args()

    config = read_config_file(args.sample.json)
    print(config)  # For testing purposes
