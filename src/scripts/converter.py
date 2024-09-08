import argparse
import io
import shutil
from datetime import datetime
from typing import Tuple

from ics import Calendar, Event
from more_itertools import always_iterable
from openpyxl import open as open_workbook
from openpyxl.cell.read_only import EmptyCell

from utils import allowed_file


def convert(filename: str) -> io.StringIO:
    file = open_workbook(filename, read_only=True)
    # TODO: scan for OK sheet in file
    sheet = file[file.sheetnames[0]]

    calendar = Calendar()
    for date, name, location, description in parse_rows(sheet.iter_rows()):
        event = Event(name=name, location=location, description=description)
        date = list(always_iterable(date))
        event.begin = date[0]
        try:
            event.end = date[1]
        except IndexError:
            pass
        event.make_all_day()  # For now support only all day events
        calendar.events.add(event)

    file = io.StringIO()
    file.writelines(calendar.serialize_iter())
    file.seek(0)  # Put 'cursor' at the begining of the file
    return file


def parse_rows(rows):
    next(rows)  # Skip headers
    for row in rows:
        parsed_row = []
        # Skip rows with missing date or name of the event
        if isinstance(row[0], EmptyCell) or isinstance(row[1], EmptyCell):
            continue
        for cell in row:
            if not isinstance(cell, EmptyCell):
                parsed_row.append(parse_date_range(cell.value)
                                  if cell == row[0] and isinstance(cell.value, str) else cell.value)
            else:
                parsed_row.append(None)
        yield parsed_row


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
