import requests
from lxml import etree
from random import sample
from openpyxl import Workbook
from bs4 import BeautifulSoup
import argparse


def fetch_list_from_url(url):
    xml = requests.get(url).content
    root = etree.XML(xml)
    return [link.text for link in root.iter('{*}loc')]


def get_course_info(course_html):
    soup = BeautifulSoup(course_html, 'html.parser')
    title = soup.find('h1', class_='title').text
    start_date = soup.find(
        'div', class_='startdate'
    ).text if soup.find(class_='startdate') else None
    start_date = start_date.split(maxsplit=1)[1] if start_date else None
    languages = soup.find('div', class_='language-info').text
    language = languages.split(',')[0]
    duration_in_weeks = len(soup.find_all(
        'div', class_='week')
    )
    rating_tag = soup.find('div', class_='rating_text')
    if rating_tag and rating_tag.text:
        rating = rating_tag.text.split()[0]
    else:
        rating = None
    return {
        'title': title,
        'starting_date': start_date,
        'language': language,
        'duration_in_weeks': duration_in_weeks,
        'rating': rating
    }


def output_courses_info_to_xlsx(courses_info):
    excel_workbook = Workbook()
    sheet = excel_workbook.active
    sheet.title = 'Courses are from coursera.com'
    column_names = [
        'Course title',
        'Starting date',
        'Language',
        'Duration (weeks)',
        'Rating'
    ]
    excel_workbook.active.append(column_names)
    for course in courses_info:
        sheet.append([
            course['title'],
            course['starting_date'],
            course['language'],
            course['duration_in_weeks'],
            course['rating']
        ])
        return excel_workbook


def save_courses_in_excel_workbook(filepath, excel_workbook):
    excel_workbook.save(filepath)


def get_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-o', '--output', help='output path')
    return parser.parse_args()


if __name__ == '__main__':
    courses_quantity = 20
    url = 'https://www.coursera.org/sitemap~www~courses.xml'

    print('The courses are loaded from coursera.com {}'.format(url))

    courses_html = fetch_list_from_url(url)
    random_courses_urls = sample(courses_html, courses_quantity)
    print('We take random courses list \n {}'.format(random_courses_urls))

    courses_raw_pages = [
        requests.get(course_url).content for course_url in random_courses_urls]
    courses_info = [
        get_course_info(course_raw_page)
        for course_raw_page in courses_raw_pages]

    args = get_parser()
    output_path = args.output

    output_courses = output_courses_info_to_xlsx(courses_info)
    save_courses_in_excel_workbook(output_path, output_courses)
    print('Start saving courses to excel-file {}'.format(output_path))

    print('There have done')
