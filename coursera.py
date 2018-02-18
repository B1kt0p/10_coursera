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


def get_courses_list(
        top_size,
        url='https://www.coursera.org/sitemap~www~courses.xml'
):
    response = requests.get(url)
    if response.ok:
        text = response.content
        xml = etree.fromstring(text)
        url_courses_list = []
        for url in xml.getchildren()[:top_size]:
            for loc in url.getchildren():
                text = loc.text
                url_courses_list.append(text)
        return url_courses_list


def get_html(url):
    response = requests.get(url)
    if response.ok:
        return response.content


def get_course_info(page_course_html, url_course):
    parse = BeautifulSoup(page_course_html, 'html.parser')
    name_course = parse.find("h1").text
    language_course = parse.find('div', {'class': 'rc-Language'})
    if language_course:
        language_course = language_course.text
    else:
        language_course = "no data"
    begin_date_course = parse.find('div', {'class': 'startdate rc-StartDateString caption-text'}).text
    if begin_date_course == 'No Upcoming Session Available':
        begin_date_course = 'No Session'
    else:
        begin_date_course = " ".join(begin_date_course.split()[1:])
    avarage_rating = parse.find(
        'div',
        {'class': 'ratings-text bt3-hidden-xs'}
    )
    if avarage_rating:
        avarage_rating = re.findall(r'\d\.\d', avarage_rating.text)[0]
    else:
        avarage_rating = "no data"
    course_info = {
        'name': name_course,
        'language': language_course,
        'begin date': begin_date_course,
        'rating': avarage_rating,
        'url': url_course
    }
    return course_info


def create_xlsx_file():
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


def add_courses_info_to_xlsx(work_book, course_info):
    work_sheet = work_book.active
    work_sheet.append([
        course_info['name'],
        course_info['language'],
        course_info['begin date'],
        course_info['rating'],
        course_info['url']
    ])
    return work_book


def save_xlsx_file(work_book, file_name):
    work_book.save(filename=file_name)


def print_course_info(course_info):
    print("{url}:".format(url=course_info['url']))
    for key, value in course_info.items():
        print('\t\t\t\t\t{key}: {value}'.format(key=key, value=value))


if __name__ == '__main__':
    file_name = get_argv().name
    top_size = 20
    work_book = create_xlsx_file()
    try:
        courses_list = get_courses_list(top_size)
        for url_course in courses_list:
            page_course_html = get_html(url_course)
            if page_course_html:
                course_info = get_course_info(page_course_html, url_course)
                add_courses_info_to_xlsx(work_book, course_info)
                print_course_info(course_info)
            else:
                continue
        save_xlsx_file(work_book, file_name)
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
        print("Can not connect!")
