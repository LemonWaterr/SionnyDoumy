import openpyxl as xl
from openpyxl.styles import PatternFill
from datetime import date

class StockInfo():

    def __init__(self, stock_file_path, order_file_path):
        self.stock_file_path = stock_file_path
        self.order_file_path = order_file_path

        # 새주문
        self.existing_stock_codes = set()
        self.new_order_stock_codes = set()
        self.new_order_stock_codes_to_row = dict() # codes : rownumber
        self.evaluate_files()

        # 바잉리스트
        self.buying_list = dict()

    # public methods
    def make_highlighted_new_order_sheet(self):
        fill_obj = PatternFill("solid", fgColor="FFFF00")

        rows_to_highlight = self.get_new_order_existing_stocks_row_numbers()
        order_xlwb = xl.load_workbook(self.order_file_path)
        for sheet_name in order_xlwb.sheetnames:
            if sheet_name != '윈런던':
                sheet_temp = order_xlwb[sheet_name]
                order_xlwb.remove(sheet_temp)
        order_xl = order_xlwb['윈런던']

        row_count=1
        for row in order_xl.iter_rows(min_row=2, max_col=34):
            row_count += 1
            if row_count in rows_to_highlight:
                for cell in row:
                    cell.fill = fill_obj

        today = date.today()
        mmdd = today.strftime("%m%d")
        order_xlwb.save(f'{mmdd} 새주문.xlsx')

    def get_new_order_existing_stocks(self):
        return self.existing_stock_codes.intersection(self.new_order_stock_codes)

    def get_new_order_missing_stocks(self):
        return self.new_order_stock_codes - self.existing_stock_codes

    def get_new_order_existing_stocks_row_numbers(self):
        existing = self.get_new_order_existing_stocks()
        result = set()
        for code in existing:
            result.add(self.new_order_stock_codes_to_row[code])
        return result

    # private methods

    def evaluate_files(self):
        stock_xlwb = xl.load_workbook(self.stock_file_path)
        order_xlwb = xl.load_workbook(self.order_file_path)
        self.iterate_stock(stock_xlwb)
        self.iterate_order(order_xlwb)

    def iterate_stock(self, stock_xlwb):
        # fetch_existing_stock_codes
        for sheet in stock_xlwb.sheetnames:
            table = stock_xlwb[sheet]
            codes_row = table['F']
            for cell in codes_row:
                self.existing_stock_codes.add(str(cell.value))

    def iterate_order(self, order_xlwb):
        order_xl = order_xlwb['윈런던']

        columns = []
        row_count=0
        for row in order_xl.iter_rows(values_only=True):
            row_count += 1

            if row_count == 1:
                columns = row
                continue
            
            if row[0] == None:
                continue
            
            # set new_order_stock_codes
            counter = str(row[4])
            product_code = str(row[10])
            try:
                if counter[-3:] == '새주문':
                    self.new_order_stock_codes.add(product_code)
                    self.new_order_stock_codes_to_row[product_code] = row_count
            except Exception as e:
                raise Exception(f"Error: row={row_count} counter={counter} product_code={product_code} \n {str(e)}")
            
            # set buying_list


ORDER_LIST_COLUMNS = [
    '발송기한',
    '주문번호',
    '상품주문번호',
    '주문일자',
    '카운터',
    '주문상태',
    '주문메모',
    '배송메모',
    '상담메모',
    '재고',
    '상품코드',
    '원주문',
    '배송GBP',
    '수량',
    '옵션',
    '원산지',
    '수령자',
    '주문자ID',
    '전화번호',
    '핸드폰번호',
    '상품금액',
    '결제금액',
    '#',
    '주소',
    '우편번호',
    '개인통관부호',
    '배송메시지',
    '브랜드',
    '품목',
    '일반품목'
]

BUYING_LIST_COLUMNS = [
    '재고',
    '발송기한',
    '상품주문번호',
    '수령자',
    '주문일자',
    '카운터',
    '주문상태',
    '배송GBP',
    '수량',
    '옵션',
    '상품코드',
    '원주문',
    '결제금액',
    #unnecessary
    '주문번호',
    '주문메모',
    '배송메모',
    '상담메모',
    '원산지',
    '주문자ID',
    '전화번호',
    '핸드폰번호',
    '상품금액',
    '#',
    '주소',
    '우편번호',
    '개인통관부호',
    '배송메시지',
    '브랜드',
    '품목',
    '일반품목'
]


stock = 'C:\\Users\\1004w\\Downloads\\재고문서 02-06 (마감) 백업본.xlsx'
order = 'C:\\Users\\1004w\\Downloads\\오더리스트 2023-02-06.xlsx'
test = StockInfo(stock, order)
test.get_new_order_existing_stocks()