import os
import requests
from pyquery import PyQuery as pq
import secret
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


def get(url, foldername, filename):
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
        # 发送网络请求, 把结果写入到文件夹
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) '
            'Chrome/70.0.3538.110 '
            'Safari/537.36',
            'Cookie': secret.cookie_boss,
        }
        r = requests.get(url, headers=headers)
        print('status_code', r.status_code)
        with open(path, 'wb') as f:
            f.write(r.content)
            return r.content


def parse_stage_area(text):
    ps = ['轮', '资', '市']
    for p in ps:
        res = text.split(p, 1)
        if len(res) == 2:
            return res[0] + p, res[1]
    return None


def query_page(page):
    e = pq(page)
    sub_li = e('.sub-li')
    res = []
    for div in sub_li:
        form = dict()
        s = pq(div)
        # 公司的名称
        form['name'] = s('h4').text()
        # 公司的详情页url
        company_url = 'https://www.zhipin.com'
        company_url += s.find('a').eq(0).attr('href')
        form['url'] = company_url
        # 公司的融资阶段，领域
        sa_text = s('p').eq(0).text()
        stage_area = parse_stage_area(sa_text)
        if stage_area is not None:
            form['stage'] = stage_area[0]
            form['area'] = stage_area[1]
        else:
            form['stage'] = sa_text

        c = Company(form)
        # print('company', c)
        res.append(c)

    return res


def output_to_excel(companies, filename):
    wb = Workbook()
    # grab the active worksheet
    ws = wb.active
    for c in companies:
        # Rows can also be appended
        ws.append(c.to_list())
    # Save the file
    wb.save(filename)


def main():
    template = 'https://www.zhipin.com/gongsi/_zzz_c101280600/'
    template += '?page={}&expectId=d0881073f0ef312e0nx83Nm6F1c~&ka=page-{}'
    companies = []
    for i in range(1, 10):
        url = template.format(i, i)
        foldername = 'cache_boss'
        filename = '{}.html'.format(i)
        page = get(url, foldername, filename)
        c = query_page(page)
        companies.extend(c)
        # print(page.decode())

    filename = 'boss_python_company.xlsx'
    output_to_excel(companies, filename)


if __name__ == '__main__':
    main()
