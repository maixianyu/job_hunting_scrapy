import os
from pyquery import PyQuery as pq
from openpyxl import Workbook


class Company(object):
    count = 0

    def __init__(self, form):
        self.id = self.count + 1
        self.add_count()
        self.name = form.get('name', 'None')
        self.url = form.get('url', 'None')
        self.stage = form.get('stage', 'None')
        self.area = form.get('area', 'None')
        self.people = form.get('people', 'None')
        self.rate_num = form.get('rate_num', 'None')
        self.position_num = form.get('position_num', 'None')
        self.resume_rate = form.get('resume_rate', 'None')

    @classmethod
    def add_count(cls):
        cls.count += 1

    def to_list(self):
        return list(self.__dict__.values())

    def __repr__(self):
        def f(k, v):
            return str(k) + ':' + str(v)
        res = '\n'.join([f(k, v) for k, v in self.__dict__.items()])
        return '<\n' + res + '\n>'


def get(foldername, filename):
    """
    缓存, 避免重复下载网页浪费时间
    """
    folder = foldername
    # 建立 cached 文件夹
    if not os.path.exists(folder):
        os.makedirs(folder)

    path = os.path.join(folder, filename)
    if os.path.exists(path):
        with open(path, 'rb') as f:
            s = f.read()
            return s
    else:
        print('empty file: {}, {}'.format(foldername, filename))


def parse_stage_people(text):
    res = text.split('/')
    if len(res) == 3:
        return res[0], res[1], res[2]
    else:
        return None


def parse_bottom_item(it):
    item = pq(it)
    return item('p').eq(0).text()


def query_page(page):
    e = pq(page)
    sub_li = e('.company-item')
    res = []
    for div in sub_li:
        form = dict()
        s = pq(div)
        # 公司的名称
        form['name'] = s('a').eq(1).text()
        # 公司的url
        form['url'] = s('a').eq(0).attr('href')
        # 公司的领域，融资，人数规模
        area_stage_people = s('.indus-stage').text()
        a_s_p = parse_stage_people(area_stage_people)
        print('area_stage_people', area_stage_people, a_s_p)
        if a_s_p is not None:
            form['area'] = a_s_p[0]
            form['stage'] = a_s_p[1]
            form['people'] = a_s_p[2]
        # 面试评价数，在招职位，简历处理率
        b_items = s('.bottom-item')
        form['rate_num'] = parse_bottom_item(b_items.eq(0))
        form['position_num'] = parse_bottom_item(b_items.eq(1))
        form['resume_rate'] = parse_bottom_item(b_items.eq(2))

        c = Company(form)
        # print('company', c)
        res.append(c)

    return res


def output_to_excel(companies, filename):
    wb = Workbook()
    # grab the active worksheet
    ws = wb.active
    # 添加说明栏
    ws.append(list(companies[0].__dict__.keys()))
    # 添加数据
    for c in companies:
        # Rows can also be appended
        ws.append(c.to_list())
    # Save the file
    wb.save(filename)


def main():
    companies = []
    for i in range(1, 11):
        foldername = 'cache_lagou'
        filename = '{}.html'.format(i)
        page = get(foldername, filename)
        c = query_page(page)
        companies.extend(c)
        # print(page.decode())

    filename = 'lagou_company.xlsx'
    output_to_excel(companies, filename)


if __name__ == '__main__':
    main()
