import argparse
import io
import shutil
from contextlib import suppress
from datetime import datetime
from typing import Iterator, Tuple

from ics import Calendar, Event
from more_itertools import always_iterable
from openpyxl import open as open_workbook

from utils import allowed_file

POSSIBLE_HEADERS = ['Data', 'Wydarzenie', 'Miejsce', 'Odpowiedzialny']


def convert(filename: str) -> io.StringIO:
    file = open_workbook(filename, read_only=True)
    # TODO: scan for OK sheet in file
    sheet = file[file.sheetnames[0]]

    calendar = Calendar()

    rows = sheet.iter_rows(values_only=True)
    headers_positions = get_header_positions(next(rows))
    if 'data' not in headers_positions or 'wydarzenie' not in headers_positions:
        raise ValueError('Excel is missing date and/or event columns')

    for row in parse_rows(rows, headers_positions):
        event = Event(name=row[headers_positions['wydarzenie']])
        date = list(always_iterable(row[headers_positions['data']]))
        event.begin = date[0]
        with suppress(IndexError):
            event.end = date[1]

        if 'miejsce' in headers_positions:
            event.location = row[headers_positions['miejsce']]
        if 'odpowiedzialny' in headers_positions:
            event.description = row[headers_positions['odpowiedzialny']]

        event.make_all_day()  # For now support only all day events
        calendar.events.add(event)

    file = io.StringIO()
    file.writelines(calendar.serialize_iter())
    file.seek(0)  # Put 'cursor' at the begining of the file
    return file


def parse_rows(rows: Iterator[tuple], header_positions):
    for row in rows:
        parsed_row = []
        # Skip rows with missing date or name of the event
        if not row[header_positions['data']] or not row[header_positions['wydarzenie']]:
            continue
        for index, val in enumerate(row):
            parsed_row.append(parse_date_range(val)
                              if index == header_positions['data'] and isinstance(val, str) else val)
        yield parsed_row


def get_header_positions(headers: tuple[str]) -> dict[str, int]:
    positions = {}
    headers = list(map(lambda x: x.lower() if isinstance(x, str) else x, headers))
    for possible_header in POSSIBLE_HEADERS:
        with suppress(ValueError):
            possible_header = possible_header.lower()
            positions[possible_header] = headers.index(possible_header)
    return positions


def parse_date_range(value: str) -> Tuple[datetime]:
    start, end = value.split('-')
    end = list(map(int, end.split('.')))
    if '.' in start:
        start = list(map(int, start.split('.')))
        start_date = datetime(end[2], start[1], start[0])
    else:
        start_date = datetime(end[2], end[1], int(start))
    return (start_date, datetime(end[2], end[1], end[0]))


def main():
    parser = argparse.ArgumentParser(
        prog="xls2ics",
        description="This tool allows for converting (for now, only specific) Excel files to ICS files"
    )
    parser.add_argument('file', type=str, help="Excel file (.xls, .xlsx)")
    args = parser.parse_args()

    assert allowed_file(args.file), 'Unsupported file'

    file = convert(args.file)
    with open('output.ics', 'w') as output_file:
        shutil.copyfileobj(file, output_file)
    file.close()
