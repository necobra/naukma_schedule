import argparse

from schedule import Schedule


def parse_arguments():
    """
    Function to parse command-line arguments for the script.

    Returns:
        argparse.Namespace: Parsed arguments.
    """
    parser = argparse.ArgumentParser(description='Process university schedules from Excel files.')

    parser.add_argument('files', metavar='file', type=str, nargs='+',
                        help='Excel file(s) to process')
    parser.add_argument('-o', '--output', metavar='output_file', type=str,
                        help='Output JSON file to save the processed schedule')

    return parser.parse_args()


if __name__ == "__main__":
    # Parse command-line arguments
    args = parse_arguments()

    schedule = Schedule(args.files)

    if args.output:
        schedule.save_to_json(args.output)
    else:
        schedule.save_to_json()
