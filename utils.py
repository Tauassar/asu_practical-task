import datetime
import logging
import re

logger = logging.getLogger(__name__)


def check_iin(iin):
    """
    Function to check whether the input iin is correct or not
    :param iin:str: actual iin
    :return:bool True if iin is correct, false otherwise
    """
    try:
        if len(iin) != 12:
            raise IndexError
        # retrieve last number to check it with control sum
        a12 = int(iin[-1])
        # iin numbers to be used in control sum calculations
        numbers = iin[:-1]
        numbers_sum = 0
        for i in range(len(numbers)):
            numbers_sum += int(numbers[i]) * (i + 1)
        control_num = numbers_sum % 11
        if control_num == 10:
            weights = [3, 4, 5, 6, 7, 8, 9, 10, 11, 1, 2]
            numbers_sum = 0
            for i in range(len(numbers)):
                numbers_sum += int(numbers[i]) * weights[i]
            control_num = numbers_sum % 11
            if 0 > control_num > 9:
                return False
        return True if control_num == a12 else False
    except IndexError:
        logger.warning('Length of IIN/BIN is incorrect')
        return False


def date_check(date):
    """
    Check date in first 4 digits of BIN
    :param date:str
    :return:bool True if not earlier than independence year
    and not later than today's day, false otherwise
    """
    # independence date
    initial_date = '0131'
    independence_date = datetime.datetime.strptime(initial_date, '%m%y')
    current_date = datetime.datetime.now()
    try:
        registration_date = datetime.datetime.strptime(date, '%m%y')
    except ValueError as e:
        logger.warning('Incorrect month for BIN\n {0}'.format(e))
        return False

    if independence_date > registration_date > current_date:
        logger.info('BIN date is incorrect')
        return False
    return True


def check_bin(bin):
    """
    Check whether bin is correct or not
    :param bin:str actual bin number
    :return:bool True if correct, false otherwise
    """
    # registration date check
    if not date_check(bin[0:4]): return False
    logger.debug('date check passed')
    if int(bin[5]) not in [4, 5, 6]:
        logger.info('second part test failed')
        return False
    if int(bin[6]) not in [0, 1, 2, 3]:
        logger.info('third part test failed')
        return False
    logger.debug('third part test failed')

    # no check for 4th part applied,
    # since there is no data about registration number

    # According to https://adilet.zan.kz/rus/docs/P030000565_
    # bin checked the same way as iin
    return check_iin(bin)


def check_single_org_form(org_name, org_form):
    """
    move organisational form to beginning
    :param org_name:str name of the organisation
    :param org_form: TOO ИП etc.
    :return:str converted form
    """
    if re.search(' '+org_form+' ', org_name):
        org_name = re.sub(' '+org_form, '', org_name)
        return '{0} {1}'.format(org_form, org_name)
    return None


def org_form_to_start(org_name):
    """
    driver for check_single_org_form,
    sends organisational form to the dedicated function
    :param org_name:str name of organisation
    :return:str proceeded string
    """
    org_forms = ['ТОО', 'ИП', 'ПК', 'АО', 'ООО', 'ЧП',
                 'Частный нотариус', 'ТДО', 'КТ']
    for form in org_forms:
        out = check_single_org_form(org_name, form)
        if out:
            return out
    # in case organisation has no org_form
    return org_name


def fl_or_ul():
    pass
