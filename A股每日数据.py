import requests
import xlwt


class Hexun(object):
    def __init__(self):
        # 构建url
        self.url = 'http://quote.tool.hexun.com/hqzx/stocktype.aspx?columnid=5500&type_code=Y0153&sorttype=3&updown=up&page=1&count=5000'
        # 构建请求头
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/61.0.3163.100 Safari/537.36'
        }

    def get_page(self):
        # 发送请求，获取响应
        response = requests.get(self.url, headers=self.headers)
        i = response.content[14:-56].decode('gbk')
        # return '[' + response.content[14:-54].decode('gbk').encode('utf8').decode('utf8')
        return i

    def pase_data(self, str_data):
        '''数据筛选'''
        list_data = []
        node_list_data = str_data.split('],[')
        for i, node in enumerate(node_list_data):
            temp_list = [node.split(',')]
            temp_list[0][0] = temp_list[0][0][1: -1]
            temp_list[0][1] = temp_list[0][1][1: -1]
            temp_list[0].append('http://vol.stock.hexun.com/%s.shtm' % (temp_list[0][0]))
            temp_list[0].append('http://guba.hexun.com/%s,guba.html' % (temp_list[0][0]))
            list_data.append(temp_list[0])
        self.print_data(list_data)
        # print(len(list_data))
        return list_data

    def save_data(self, list_data):
        '''传入列表数据，保存为表格'''
        # 创建对象
        workbook = xlwt.Workbook()
        # 创建工作表
        sheet = workbook.add_sheet('A股详情', cell_overwrite_ok=True)
        sheet.write(0, 0, '代码')
        sheet.write(0, 1, '名称')
        sheet.write(0, 2, '最新价')
        sheet.write(0, 3, '涨跌幅')
        sheet.write(0, 4, '昨收')
        sheet.write(0, 5, '今天')
        sheet.write(0, 6, '最高')
        sheet.write(0, 7, '最低')
        sheet.write(0, 8, '成交量')
        sheet.write(0, 9, '成交额')
        sheet.write(0, 10, '振幅')
        sheet.write(0, 11, '量比')
        sheet.write(0, 12, '资金图')
        sheet.write(0, 13, '和讯股吧')
        for row, temp_list in enumerate(list_data):
            for column, node in enumerate(temp_list):
                sheet.write(row + 1, column, node)
        workbook.save('./A股今日数据.xls')

    def print_data(self, data):
        '''打印测试'''
        print(type(data))
        print(data)

    def run(self):
        # 发送请求获取响应
        str_data = self.get_page()
        # 数据筛选
        list_data = self.pase_data(str_data)
        # 数据存储
        self.save_data(list_data)


if __name__ == '__main__':
    he = Hexun()
    he.run()
