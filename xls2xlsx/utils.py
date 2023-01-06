from fractions import Fraction
from openpyxl.utils.datetime import from_excel
# broken! from openpyxl.styles.numbers import is_date_format
from openpyxl.styles.numbers import STRIP_RE
from datetime import datetime, date, timedelta
from datetime import time as tm
import re

TRACE=False

def is_date_format(fmt):            # The one in openpyxl doesn't work properly!
    if fmt is None:
        return False
    fmt = fmt.split(";")[0] # only look at the first format
    fmt = STRIP_RE.sub("", fmt) # ignore some formats
    return re.search(r"(?<!\\)[dmhysDMHYS]", fmt) is not None

def perform_number_format(value, number_format):
    """This is a half-baked attempt at formatting the given
    value using the given Excel number_format.  This is used by
    the tests to match values.  Handled is many of the formats for
    numbers (int/float), datetime, date, time, and timedelta."""

    if number_format == 'General' or isinstance(value, str):
        return value
    if number_format == '@':
        return str(value)
    grabit = []
    def grab_escapes(number_format):
        nonlocal grabit
        def sub_grabit(m):
            i = len(grabit)
            grabit.append(m.group(1))
            return f'{{{i}}}'
        nf = re.sub(r'\\(.)', sub_grabit, number_format)
        nf = re.sub(r'"([^"]*)"', sub_grabit, nf)
        nf = re.sub(r'\[(hh|h|mm|m|ss|s)\]', r'<\1>', nf)   # So we don't match the next rule with [h]
        nf = re.sub(r'\[[^\]]+\]', '', nf)   # Remove [Blue], [$-F800], etc
        nf = re.sub(r'<(hh|h|mm|m|ss|s)>', r'[\1]', nf)   # Put back the [h] etc
        return nf

    def restore_escapes(nf):
        nonlocal grabit
        if len(grabit):
            nf = nf.format(*grabit)       # Put escaped chars back in
        return nf

    if TRACE:
        print(f'perform_number_format({value}, {number_format})')
    if (isinstance(value, int) or isinstance(value, float)) and is_date_format(number_format):
        if '[h' in number_format or '[m' in number_format or '[s' in number_format:
            value = timedelta(days=value)
        else:
            value = from_excel(value)
    if isinstance(value, int) or isinstance(value, float):
        # Note: This is NOT a full implementation of Excel int/float number formatting!
        format_split = number_format.split(';')
        number_format = format_split[0]
        prefix = ''
        suffix = ''
        if value < 0 and len(format_split) >= 2:
            number_format = format_split[1]
            value = abs(value)
        elif value == 0 and len(format_split) >= 3:
            number_format = format_split[2]
        if not number_format:
            return ''
        nf = grab_escapes(number_format)
        fmt = 'f'
        if isinstance(value, int):
            fmt = 'd'
        if '%' in nf:
            fmt = '%'
        elif 'E' in nf:
            fmt = 'E'
            nf = re.sub(r'E[+0#?]+', 'E', nf)
        comma = ''
        pound = ''
        c_ndx = nf.find(',')
        d_ndx = nf.find('.')
        p_ndx = nf.find('#')
        if c_ndx >= 0:
            if d_ndx >=0 and c_ndx > d_ndx:
                while c_ndx < len(nf):
                    value /= 100
                    c_ndx = nf.find(',', c_ndx+1)
            else:
                comma = ','
        places = ''
        if d_ndx >= 0:
            if p_ndx > d_ndx:
                pound = '#'
                nf = nf.replace('#', '0')
            places = f'.{nf[d_ndx+1:].count("0")}'
            if fmt == 'd':
                value = float(value)
                fmt = 'f'
        elif fmt == 'd':
            zeros = nf.count('0')
            if zeros:
                fmt = f'0{zeros}' + fmt
        else:
            places = '.0'
        nf = re.sub(r'_.', ' ', nf)
        nf = nf.replace('*', '')        # We can't really do this one
        m = re.match(r'((?:[^0#.E%,?*]*{\d+}[^0#.E%,?*]*)|(?:[^0#.E%,?*]*))[0#.E%,?*]+(.*[/][0-9?#]+)?(.*)$', nf)
        prefix = restore_escapes(m.group(1))
        fraction = m.group(2)
        suffix = restore_escapes(m.group(3))
        if fraction:
            s_ndx = fraction.find('/')
            suf = restore_escapes(fraction[:s_ndx]).replace('?', '').replace('#', '').replace('0', '')
            fraction = fraction[s_ndx+1:]
            if fraction.isdigit():
                ival = int(value)
                value -= ival
                suf += f'{value//int(fraction)}/{fraction}'
                value = ival
            if fraction[0] != '?' or float(int(value)) != value:
                ival = int(value)
                value -= ival
                fr = Fraction.from_float(value).limit_denominator(10**(len(fraction))-1)
                suf += f'{fr.numerator}/{fr.denominator}'
            suffix = suf + suffix

        py_format = f'{prefix}{{0:{pound}{comma}{places}{fmt}}}{suffix}'
        value = py_format.format(value)
        if TRACE:
            print(f'perform_number_format: using {py_format} to produce {value}')
        return value

    number_format = number_format.split(';')[0]
    if isinstance(value, tm):
        value = datetime(1, 1, 1, value.hour, value.minute, value.second)
    elif isinstance(value, date) and not isinstance(value, datetime):
        value = datetime(value.year, value.month, value.day)
    if isinstance(value, datetime) and \
      ('[h' in number_format or '[m' in number_format or '[s' in number_format):
        value = timedelta(hours=value.hour, minutes=value.minute, 
                          seconds=value.second + value.microsecond / 1000000.0)
    if isinstance(value, timedelta):
        total_seconds = int(value.total_seconds())
        hours = total_seconds // 3600
        total_minutes = total_seconds // 60
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        nf = grab_escapes(number_format)
        nf = nf.replace('[hh]', f'{hours:02d}').replace('[mm]', f'{total_minutes:02d}'). \
                replace('[ss]', f'{total_seconds:02d}').replace('[h]', str(hours)). \
                replace('[m]', str(total_minutes)).replace('[s]', str(total_seconds)). \
                replace('mm', f'{minutes:02d}').replace('ss', f'{seconds:02d}'). \
                replace('m', str(minutes)).replace('s', str(seconds))
        nf = restore_escapes(nf)
        value = nf
        if TRACE:
            print(f'perform_number_format: timedelta produced {value} (grabit = {grabit})')
    if isinstance(value, datetime):
        if value.microsecond >= 500000:        # Round up 999999 ms to the next second
            value = value.replace(microsecond=0) + timedelta(seconds=1)
        fmt = grab_escapes(number_format)
        fmt = fmt.replace('yyyy', '%Y').replace('yy', '%y').replace('dddd', '%A').replace('ddd', '%a'). \
                replace('dd', '%D').replace('mmmm', '%B').replace('mmm', '%b').replace('AM/PM', '%p'). \
                replace('ss', '%S')
        h_ndx = fmt.find('h')
        if '%p' in fmt:
            fmt = fmt.replace('hh', '%I')
        else:
            fmt = fmt.replace('hh', '%H')
        # Now let's handle the hard ones: mm, m, d, h, a/p
        ap_ndx = fmt.find('a/p')
        if ap_ndx >= 0:
            fmt = fmt.replace('a/p', '%p')
        while True:
            m_ndx = fmt.find('mm')
            if m_ndx < 0:
                break
            if h_ndx >= 0 and m_ndx > h_ndx: # it's minutes
                fmt = fmt[:m_ndx] + '%M' + fmt[m_ndx+2:]
                continue
            fmt = fmt[:m_ndx] + '%X' + fmt[m_ndx+2:]    # it's months (corrected below)
        while True:
            m_ndx = fmt.find('m')
            if m_ndx < 0:
                break
            if h_ndx >= 0 and m_ndx > h_ndx: # it's minutes
                fmt = fmt[:m_ndx] + str(value.minute) + fmt[m_ndx+1:]
                continue
            fmt = fmt[:m_ndx] + str(value.month) + fmt[m_ndx+1:]    # it's months
        d_ndx = fmt.find('d')
        if d_ndx >= 0:
            fmt = fmt.replace('d', str(value.day))
        if h_ndx >= 0:
            if '%p' in fmt:
                hour = value.hour
                if hour > 12:
                    hour -= 12
                if hour == 0:
                    hour = 12
                fmt = fmt.replace('h', str(hour))
            else:
                fmt = fmt.replace('h', str(value.hour))

        fmt = fmt.replace('%D', '%d').replace('%X', '%m')
        fmt = restore_escapes(fmt)
        value = value.strftime(fmt)
        if ap_ndx >= 0:
            value = value.replace('AM', 'a').replace('PM', 'p')
        if TRACE:
            print(f'perform_number_format: using {fmt} to produce {value} (grabit={grabit})')
    return value
