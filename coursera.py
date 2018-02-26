import requests
from openpyxl import Workbook
from lxml import etree
from bs4 import BeautifulSoup
import re
import argparse


def get_argv():
    parser = argparse.ArgumentParser(
        description='infirmation about courses'
    )
    parser.add_argument(
        '--name',
        '-n',
        default='coursera.xlsx',
        help='name exel file'
    )
    return parser.parse_args()


def get_courses_url_list(
        top_size,
        xml_data
):
    xml = etree.fromstring(xml_data)
    url_courses_list = []
    for url in xml.getchildren()[:top_size]:
        for loc in url.getchildren():
            text = loc.text
            url_courses_list.append(text)
    return url_courses_list


def fetch_content(url):
    try:
        response = requests.get(url)
        if response.ok:
            return response.content
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
        return None


def get_course_info(page_course_html, url_course):
    soup = BeautifulSoup(page_course_html, 'html.parser')
    name_course = soup.find('h1').text
    language_course = soup.find('div', {'class': 'rc-Language'})
    if language_course:
        language_course = language_course.text
    else:
        language_course = None
    begin_date_course = soup.find(
        'div',
        {'class': 'startdate rc-StartDateString caption-text'}
    )
    if begin_date_course.text == 'No Upcoming Session Available':
        begin_date_course = None
    else:
        begin_date_course = ''.join(begin_date_course.text.split()[1:])
    avarage_rating = soup.find(
        'div',
        {'class': 'ratings-text bt3-hidden-xs'}
    )
    if avarage_rating:
        avarage_rating = re.findall(r'\d\.\d', avarage_rating.text)[0]
    else:
        avarage_rating = None
    course_info = {
        'name': name_course,
        'language': language_course,
        'begin date': begin_date_course,
        'rating': avarage_rating,
        'url': url_course
    }
    return course_info


def create_work_book():
    work_book = Workbook()
    work_sheet = work_book.active
    title = [
        'Name',
        'Language',
        'Begin date',
        'Avarage_rating',
        'url'
    ]
    work_sheet.append(title)
    return work_book


def add_courses_info_to_work_sheet(work_book, course_info):
    work_sheet = work_book.active
    added_row = ['no data'
                 if course_info[key] is None
                 else course_info[key]
                 for key in course_info
                 ]
    work_sheet.append(added_row)
    return work_book


def save_xlsx_file(work_book, file_name):
    work_book.save(file_name)


if __name__ == '__main__':
    file_name = get_argv().name
    top_size = 20
    url = 'https://www.coursera.org/sitemap~www~courses.xml'
    work_book = create_work_book()
    xml_data = fetch_content(url)
    if xml_data:
        courses_url_list = get_courses_url_list(top_size, xml_data)
        for url_course in courses_url_list:
            page_course_html = fetch_content(url_course)
            if page_course_html:
                course_info = get_course_info(page_course_html, url_course)
                add_courses_info_to_work_sheet(work_book, course_info)
            else:
                print('Can not open {}!'.format(url_course))
        save_xlsx_file(work_book, file_name)
        print('File {} saved.'.format(file_name))
    else:
        print('Can not connect!')
