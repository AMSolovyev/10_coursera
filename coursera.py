import requests
from lxml import etree
from random import sample
from openpyxl import Workbook
from bs4 import BeautifulSoup


def get_courses(url):
    xml = requests.get(url).content
    return xml


def get_data(xml):
    root = etree.XML(xml)
    return [link.text for link in root.iter('{*}loc')]


def get_course_info(course_all_info):
    soup = BeautifulSoup(course_all_info, 'html.parser')
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
    return {'title': title,
            'starting_date': start_date,
            'language': language,
            'duration_in_weeks': duration_in_weeks,
            'rating': rating}


def output_courses_info_to_xlsx(courses_info):
    excel_workbook = Workbook()
    sheet = excel_workbook.active
    sheet.title = 'Courses are from coursera.com'
    column_names = [
        'Course title', 'Starting date',
        'Language', 'Duration (weeks)',
        'Rating'
    ]
    excel_workbook.active.append(column_names)
    for course in courses_info:
        sheet.append([
            course['title'], course['starting_date'],
            course['language'], course['duration_in_weeks'],
            course['rating']
        ])
        return excel_workbook


def save_courses_in_excel_workbook(filepath, excel_workbook):
    excel_workbook.save(filepath)


if __name__ == '__main__':
    courses_quality = 20
    excel_file_name = 'courses.xlsx'
    url = 'https://www.coursera.org/sitemap~www~courses.xml'
    print('The courses are loaded from coursera.com {}'.format(url))

    courses = get_courses(url)
    all_courses = get_data(courses)
    random_courses = sample(all_courses, courses_quality)
    print('We take random courses list \n {}'.format(random_courses))

    courses_raw_pages = [
        requests.get(course_url).content for course_url in random_courses]
    courses_info = [
        get_course_info(course_raw_page)
        for course_raw_page in courses_raw_pages]

    output_courses = output_courses_info_to_xlsx(courses_info)
    save_courses_in_excel_workbook(excel_file_name, output_courses)
    print('Start saving courses to excel-file {}'.format(excel_file_name))

    print('There have done')
