import StockInfo

stock = 'C:\\Users\\1004w\\Downloads\\재고문서 02-06 (마감) 백업본.xlsx'
order = 'C:\\Users\\1004w\\Downloads\\오더리스트 2023-02-06.xlsx'
test = StockInfo(stock, order)
test.execute()
#test.get_new_order_existing_stocks()