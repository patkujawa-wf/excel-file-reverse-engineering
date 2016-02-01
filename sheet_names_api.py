# coding=utf-8
"""

> echo 'xlsx/SOX Controls Testing Template.xlsx' | python sheet_names_.py --verbosity=info
> for fname in **/*.xlsx; do time echo $fname | python sheet_names_.py; done
> \ls -1 **/*.xls? | python sheet_names_.py --verbosity=info
> find . -print
> \ls -d -1 **/*.*
"""
import argparse
import logging as logging
import sys
import time
import warnings

_logger = logging.getLogger(__name__)


def _out(msg, *values):
    if values:
        if '%' in msg:
            msg = msg % values
        elif '{' in msg:
            msg = msg.format(values)
    print(msg)


killers = [
    'xlsx/SOX Controls Testing Template.xlsx',  # 31,728,240 bytes
    'xlsx/GAAP Check List Passworded <OBFUSCATED>.xlsx', # ERROR:sheet_names_api:BadZipfile('File is not a zip file',)
]


def _loglevel_from_str(s):
    """

    :param s:
    :type s: str
    """
    d = dict(
        CRITICAL=logging.CRITICAL,
        DEBUG=logging.DEBUG,
        ERROR=logging.ERROR,
        FATAL=logging.FATAL,
        INFO=logging.INFO,
        NOTSET=logging.NOTSET,
        WARN=logging.WARN,
        WARNING=logging.WARNING,
    )
    return d.get(s.upper(), logging.NOTSET)


def get_options():
    parser = argparse.ArgumentParser(
        add_help=True,
        description='Use --verbosity=0 to quiet',
    )
    parser.add_argument(
        '--verbosity', dest='loglevel', default=logging.DEBUG, type=_loglevel_from_str
    )
    return parser.parse_args()


def map_smaller_file(func, filepath, suffix='.xlsx', truncated_size=2 ** 20):
    """

    :param func: (filepath) => object
    :type func: Function
    :param filepath:
    :type filepath: str
    :param suffix: openpyxl, for instance, does a file extension check
    :type suffix: str
    :param truncated_size:
    :type truncated_size:
    :return:
    :rtype:
    """
    import io
    import tempfile
    # NOTE: can't return temp file path from here because it is destroyed, so instead we execute a user-supplied function upon the file while it's in scope
    dest = tempfile.NamedTemporaryFile(suffix=suffix)
    dest_path = dest.name
    with io.open(dest_path, 'wb') as dest_stream:
        with io.open(filepath, 'rb') as of:
            dest_stream.write(of.read(truncated_size))
    return func(dest_path)
    # openpyxl.load_workbook(dest_path, read_only=True)
    # BadZipfile: File is not a zip file


def main(get_sheet_names, lines=None):
    opts = get_options()
    # warnings.filterwarnings('ignore')  # UserWarning: Discarded range with reserved name
    logging.basicConfig(level=opts.loglevel)
    slow_threshold = 0.500  # seconds
    slow_files = []  # will hold tuples
    for line in lines or sys.stdin:
        filepath = line.strip()
        try:
            _logger.debug('%s >', filepath)
            s = time.time()
            sheet_names = get_sheet_names(filepath)
            seconds_taken = time.time() - s
            if seconds_taken > slow_threshold:
                slow_item = (seconds_taken, filepath)
                slow_files.append(slow_item)
            _logger.debug('%r seconds', seconds_taken)
            _logger.debug(sheet_names)
        except Exception as e:
            _logger.exception('%r for filepath %r', e, filepath)

    if slow_files:
        slow_files.sort()
        from pprint import pformat
        _logger.info('Slow (took longer than %s seconds) files:\n%s', slow_threshold, pformat(slow_files))
