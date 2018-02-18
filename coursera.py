import requests
from openpyxl import Workbook
from lxml import etree
from bs4 import BeautifulSoup
import re


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


def get_course_info(course_slug):
    response = requests.get(course_slug)
    if response.ok:
        parse = BeautifulSoup(response.content, 'html.parser')
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
        course_inf = {
            'name': name_course,
            'language': language_course,
            'begin date': begin_date_course,
            'rating': avarage_rating
        }
        return course_inf


def output_courses_info_to_xlsx(file_name, top_size):
    url_courses_list = get_courses_list(top_size)
    if url_courses_list:
        work_book = Workbook()
        work_sheet = work_book.active
        for url_courses in url_courses_list:
            course_info = get_course_info(url_courses)
            if course_info:
                work_sheet.append(list(course_info.values()))
                print("{url}:".format(url=url_courses))
                for key, value in course_info.items():
                    print('\t\t\t\t\t{key}: {value}'.format(
                        key=key,
                        value=value
                    ))
            else:
                print ('{} - can not connect.'.format(url_courses))
                continue
        work_book.save(filename=file_name)
        print ('Information is written to a file {}'.format(file_name))


if __name__ == '__main__':
    try:
        top_size = 20
        file_name = "coursera.xlsx"
        output_courses_info_to_xlsx(file_name, top_size)
    except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
        print("Can not connect!")
